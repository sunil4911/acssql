# Extracts the reusable "letterhead" portion (firm-info sidebar, divider line,
# and FDCPA disclosure footer) out of a source PDF page, discarding whatever
# one-off letter body happens to be on that particular page.
#
# The source PDF (e.g. LH.pdf) is expected to be a single page with an EMPTY
# /Resources dictionary (no fonts/images/XObjects) - i.e. all of its content is
# pure vector fill/stroke path operators. This is what "Print to PDF" produces
# when text has been converted to outline paths. Each glyph/stroke is emitted
# as its own independent, self-contained "q ... Q" block (no cross-block state
# leakage), which lets us classify and keep/drop whole blocks by where their
# bounding box lands on the page, rather than needing OCR or font/text data.
#
# Classification rule (tuned for DOC\LH.pdf's specific layout - a right-hand
# sidebar column starting around x=446pt, and an optional bottom disclosure
# footer within ~60pt of the page bottom): a block is kept if its horizontal
# position is in the sidebar column, and optionally if its vertical position is
# in the footer band when -IncludeFooter is set. Everything else (the
# recipient/body/signature text on the left) is dropped.
#
# By default the footer is NOT included - clsPDF.SetLetterheadFooter draws the
# bottom disclosure as real, variable text instead.
#
# Run this again (adjusting $SidebarMinX / $FooterMaxY if the source letterhead
# layout ever changes) whenever DOC\LH.pdf is replaced, then clsPDF.SetLetterhead
# picks up the regenerated DOC\LH_Content.txt automatically.

param(
    [string]$SrcPdf = 'C:\ACSSQL\DOC\LH.pdf',
    [string]$OutFile = 'C:\ACSSQL\DOC\LH_Content.txt',
    [double]$SidebarMinX = 430,
    [double]$FooterMaxY = 60,
    [switch]$IncludeFooter
)

Add-Type -AssemblyName System.IO.Compression

$bytes = [System.IO.File]::ReadAllBytes($SrcPdf)
$text = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)

function DecompressStream([int]$num) {
    $pattern = "(?s)(?<![0-9])$num\s+0\s+obj.*?stream\r?\n"
    $m = [regex]::Match($text, $pattern)
    if (-not $m.Success) { throw "stream object $num not found" }
    $startIdx = $m.Index + $m.Length
    $endMatch = [regex]::new('endstream').Match($text, $startIdx)
    $endIdx = $endMatch.Index
    while ($text[$endIdx-1] -eq "`n" -or $text[$endIdx-1] -eq "`r") { $endIdx-- }
    $len = $endIdx - $startIdx
    $raw = New-Object byte[] $len
    [Array]::Copy($bytes, $startIdx, $raw, 0, $len)
    $ms = New-Object System.IO.MemoryStream(,$raw)
    $ms.ReadByte(); $ms.ReadByte()  # skip 2-byte zlib header
    $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Decompress)
    $out = New-Object System.IO.MemoryStream
    $ds.CopyTo($out)
    return [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetString($out.ToArray())
}

# --- Step 1: find the page's /Contents object number(s) and decompress+concat them ---
$pageMatch = [regex]::Match($text, '/Contents\s*(\[[^\]]*\]|\d+\s+0\s+R)')
if (-not $pageMatch.Success) { throw "could not find a /Contents entry" }
$contentsSpec = $pageMatch.Groups[1].Value
$contentNums = [regex]::Matches($contentsSpec, '(\d+)\s+0\s+R') | ForEach-Object { [int]$_.Groups[1].Value }

$sb = New-Object System.Text.StringBuilder
foreach ($n in $contentNums) { [void]$sb.AppendLine((DecompressStream $n)) }
$rawFull = $sb.ToString()

# --- Normalize operand counts per line ---
# "Print to PDF" output has been observed to leave stray extra numeric operands
# before some operators (e.g. "120 156 0.75 0 0 -0.75 0 792 cm" - 8 numbers for
# an operator that only takes 6). Lenient viewers (Adobe, Chrome) just discard
# excess operands and keep going, but Windows' own PDF renderer aborts the
# *entire* content stream on the mismatch, silently rendering a blank page. Per
# operator, keep only the last N numeric operands it actually needs.
$operandCounts = New-Object 'System.Collections.Generic.Dictionary[string,int]'  # case-sensitive (ordinal) keys
foreach ($kv in @(
    @('cm',6), @('m',2), @('l',2), @('c',6), @('v',4), @('y',4), @('re',4),
    @('rg',3), @('RG',3), @('g',1), @('G',1), @('k',4), @('K',4),
    @('w',1), @('J',1), @('j',1), @('M',1)
)) { $operandCounts[$kv[0]] = $kv[1] }
$zeroOperandOps = @('h','f','f*','S','s','n','q','Q','W','W*','B','B*','b','b*')

$normLines = New-Object System.Collections.Generic.List[string]
foreach ($rawLine in ($rawFull -split "`r?`n")) {
    $t = $rawLine.Trim()
    if ($t -eq '') { $normLines.Add($rawLine); continue }
    $toks = $t -split '\s+'
    $op = $toks[$toks.Length-1]
    if ($zeroOperandOps -ccontains $op) {
        $normLines.Add($op)
    }
    elseif ($operandCounts.ContainsKey($op)) {
        $need = $operandCounts[$op]
        $numToks = $toks[0..($toks.Length-2)]
        if ($numToks.Length -gt $need) {
            $numToks = $numToks[($numToks.Length-$need)..($numToks.Length-1)]
        }
        $normLines.Add((($numToks + $op) -join ' '))
    }
    else {
        $normLines.Add($rawLine)
    }
}
$full = ($normLines -join "`n") + "`n"

# --- Step 2: split into the page-wide preamble + independent q...Q blocks ---
$firstQIdx = $full.IndexOf("q`n")
$preamble = $full.Substring(0, $firstQIdx).Trim()
$rest = $full.Substring($firstQIdx)

$markers = [regex]::Matches($rest, '(?m)^(q|Q)\r?$')
$chunks = New-Object System.Collections.Generic.List[string]
$depth = 0; $chunkStart = -1
foreach ($m in $markers) {
    if ($m.Groups[1].Value -ceq 'q') {
        if ($depth -eq 0) { $chunkStart = $m.Index }
        $depth++
    } else {
        $depth--
        if ($depth -eq 0) { $chunks.Add($rest.Substring($chunkStart, $m.Index + $m.Length - $chunkStart)) }
    }
}

# --- Step 3: compute each block's device-space bounding box ---
function MakeMatrix([double[]]$nums) {
    $m = New-Object 'double[]' 6
    for ($i=0;$i -lt 6;$i++) { $m[$i] = $nums[$i] }
    return ,$m
}
function MatMul([double[]]$m1, [double[]]$m2) {
    $a1=$m1[0]; $b1=$m1[1]; $c1=$m1[2]; $d1=$m1[3]; $e1=$m1[4]; $f1=$m1[5]
    $a2=$m2[0]; $b2=$m2[1]; $c2=$m2[2]; $d2=$m2[3]; $e2=$m2[4]; $f2=$m2[5]
    $out = New-Object 'double[]' 6
    $out[0] = $a1*$a2 + $b1*$c2
    $out[1] = $a1*$b2 + $b1*$d2
    $out[2] = $c1*$a2 + $d1*$c2
    $out[3] = $c1*$b2 + $d1*$d2
    $out[4] = $e1*$a2 + $f1*$c2 + $e2
    $out[5] = $e1*$b2 + $f1*$d2 + $f2
    return ,$out
}

$preambleToks = $preamble -split '\s+'
$outerM = MakeMatrix ([double[]]($preambleToks[($preambleToks.Length-7)..($preambleToks.Length-2)] | ForEach-Object { [double]$_ }))

$kept = New-Object System.Collections.Generic.List[string]
$keptCount = 0; $droppedCount = 0

for ($ci = 0; $ci -lt $chunks.Count; $ci++) {
    $chunk = $chunks[$ci]
    $lines = $chunk -split "`r?`n"
    $curM = $outerM
    $minX = [double]::PositiveInfinity; $minY = [double]::PositiveInfinity
    $maxX = [double]::NegativeInfinity; $maxY = [double]::NegativeInfinity
    $hasPoint = $false

    foreach ($line in $lines) {
        $t = $line.Trim()
        if ($t -eq '' -or $t -ceq 'q' -or $t -ceq 'Q' -or $t -ceq 'h' -or $t -ceq 'f' -or $t -ceq 'f*' -or $t -ceq 'S') { continue }
        $toks = $t -split '\s+'
        $op = $toks[$toks.Length-1]
        $numToks = $toks[0..($toks.Length-2)]
        $nt = $numToks.Length

        if ($op -ceq 'cm') {
            $last6 = $numToks[($nt-6)..($nt-1)]
            $nm = MakeMatrix ([double[]]($last6 | ForEach-Object { [double]$_ }))
            $curM = MatMul $nm $curM
        }
        elseif ($op -ceq 'm' -or $op -ceq 'l') {
            $x = [double]$numToks[$nt-2]; $y = [double]$numToks[$nt-1]
            $a=$curM[0]; $b=$curM[1]; $c=$curM[2]; $d=$curM[3]; $e=$curM[4]; $f=$curM[5]
            $px = $a*$x + $c*$y + $e; $py = $b*$x + $d*$y + $f
            if ($px -lt $minX) { $minX = $px }; if ($px -gt $maxX) { $maxX = $px }
            if ($py -lt $minY) { $minY = $py }; if ($py -gt $maxY) { $maxY = $py }
            $hasPoint = $true
        }
        elseif ($op -ceq 'c') {
            $a=$curM[0]; $b=$curM[1]; $c=$curM[2]; $d=$curM[3]; $e=$curM[4]; $f=$curM[5]
            for ($k=$nt-6; $k -lt $nt; $k+=2) {
                $x = [double]$numToks[$k]; $y = [double]$numToks[$k+1]
                $px = $a*$x + $c*$y + $e; $py = $b*$x + $d*$y + $f
                if ($px -lt $minX) { $minX = $px }; if ($px -gt $maxX) { $maxX = $px }
                if ($py -lt $minY) { $minY = $py }; if ($py -gt $maxY) { $maxY = $py }
            }
            $hasPoint = $true
        }
        elseif ($op -ceq 're') {
            $rx=[double]$numToks[$nt-4]; $ry=[double]$numToks[$nt-3]; $rw=[double]$numToks[$nt-2]; $rh=[double]$numToks[$nt-1]
            $a=$curM[0]; $b=$curM[1]; $c=$curM[2]; $d=$curM[3]; $e=$curM[4]; $f=$curM[5]
            foreach ($corner in @(@($rx,$ry),@($rx+$rw,$ry),@($rx,$ry+$rh),@($rx+$rw,$ry+$rh))) {
                $x=$corner[0]; $y=$corner[1]
                $px = $a*$x + $c*$y + $e; $py = $b*$x + $d*$y + $f
                if ($px -lt $minX) { $minX = $px }; if ($px -gt $maxX) { $maxX = $px }
                if ($py -lt $minY) { $minY = $py }; if ($py -gt $maxY) { $maxY = $py }
            }
            $hasPoint = $true
        }
    }

    $keepThis = $false
    if ($hasPoint) {
        if ($minX -ge $SidebarMinX) { $keepThis = $true }
        elseif ($IncludeFooter -and $maxY -le $FooterMaxY) { $keepThis = $true }
    }
    if ($keepThis) { $kept.Add($chunk); $keptCount++ } else { $droppedCount++ }
}

$finalContent = $preamble + "`n" + ($kept -join "`n")
[System.IO.File]::WriteAllText($OutFile, $finalContent, [System.Text.Encoding]::GetEncoding('ISO-8859-1'))
Write-Output ("Kept $keptCount blocks, dropped $droppedCount blocks (of $($chunks.Count) total).")
Write-Output ("Wrote " + $finalContent.Length + " chars to " + $OutFile)
