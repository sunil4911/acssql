Attribute VB_Name = "Letterhead_clsPDF"
Option Explicit

' ===========================================================================
' Letterhead_clsPDF - THE single reusable letterhead template.
'
' This module is the one and only place that knows what your letterhead
' looks like (firm name, tagline, address, the "Attorneys & Counselors at
' Law" signature style, and the standard debt-collector disclosure). Every
' document module (plain letters, the Information Subpoena, future letters)
' calls INTO this module for those pieces instead of drawing its own copy.
' That means: if the letterhead ever changes (new address, new tagline,
' rename the firm, etc.), you edit the FIRM_* constants below ONE time and
' every document that uses PrintFirmHeader / PrintFirmSignatureBlock /
' PrintFDCPADisclosure picks up the change automatically the next time it
' runs - nothing else needs to be touched.
'
' Three reusable building blocks, callable from ANY module on an already-
' created clsPDF object (they do NOT call CreatePDF/SavePDF themselves, so
' you can freely mix them with your own content, multiple pages, etc.):
'
'   PrintFirmHeader PDF            - the top letterhead block + rule
'   PrintFirmSignatureBlock PDF    - name/firm/tagline/address signature block
'   PrintFDCPADisclosure PDF       - the bottom "debt collector" notice
'
' Plus one ready-to-use, all-in-one wrapper for simple single-purpose
' letters that don't need anything fancier:
'
'   PrintLetterhead "C:\temp\MyLetter.pdf", _
'       RecipientLines:="John Doe" & vbNewLine & "123 Main St." & vbNewLine & "Anytown, NY 12345", _
'       DateStr:=Format(Date, "mmmm d, yyyy"), _
'       Salutation:="Dear Mr. Doe:", _
'       Body:="First paragraph of the letter..." & vbNewLine & vbNewLine & _
'             "Second paragraph of the letter...", _
'       SignerName:="Amanda C Shaw", _
'       SignerTitle:="Attorney"
'
' Body paragraphs are separated by a blank line (vbNewLine & vbNewLine, or
' vbCrLf & vbCrLf) - each paragraph is word-wrapped and justified
' automatically; you don't need to worry about line breaks within a
' paragraph.
'
' For a document that combines a cover letter AND another document (e.g.
' the Information Subpoena) in one PDF, see
' LET10_InfoSubpoena_clsPDF.PrintSubpoenaWithCoverLetter for an example of
' calling PrintFirmHeader for page 1, PDF.NewPage, then the second
' document's own content, then PrintFirmSignatureBlock/PrintFDCPADisclosure
' again for page 2 - all sharing this same template module.
' ===========================================================================

Public Const FIRM_NAME As String = "Riehlman, Shafer & Shaw, LLC"
Public Const FIRM_TAGLINE1 As String = "ATTORNEYS & COUNSELORS AT LAW"
Public Const FIRM_TAGLINE2 As String = "Specializing in Creditors' Rights"
Public Const FIRM_ADDRESS_LINE1 As String = "Post Office Box 544"
Public Const FIRM_ADDRESS_LINE2 As String = "Tully, New York 13159-0544"
Public Const FIRM_PHONE As String = "(315) 696-6347"

' --- Sidebar geometry (inches from the page's top-left corner) -------------
' Tweak these if the sidebar needs to move/resize; everything else adapts.
Public Const SIDEBAR_DIVIDER_X As Double = 6.15
Public Const SIDEBAR_LEFT As Double = 6.35
Public Const SIDEBAR_RIGHT As Double = 8
Public Const SIDEBAR_TOP As Double = 0.6
Public Const SIDEBAR_DIVIDER_BOTTOM As Double = 10.2

' --- Shared template piece #1: the top letterhead block --------------------
' Draws the firm name, tagline(s), address, and a horizontal rule onto an
' already-created clsPDF object. Leaves the font set to Times 12 (normal)
' and the cursor just below the rule, ready for whatever comes next.
Public Sub PrintFirmHeader(ByRef PDF As clsPDF, _
    Optional ByVal FirmName As String = FIRM_NAME, _
    Optional ByVal FirmTagline1 As String = FIRM_TAGLINE1, _
    Optional ByVal FirmTagline2 As String = FIRM_TAGLINE2, _
    Optional ByVal FirmAddressLine1 As String = FIRM_ADDRESS_LINE1, _
    Optional ByVal FirmAddressLine2 As String = FIRM_ADDRESS_LINE2, _
    Optional ByVal RuleWidth As Integer = 70)

    PDF.SetFont "Times", 18, True
    PDF.PrintText FirmName, pdfCenter
    PDF.SetFont "Times", 11, True
    If Len(FirmTagline1) > 0 Then PDF.PrintText FirmTagline1, pdfCenter
    If Len(FirmTagline2) > 0 Then PDF.PrintText FirmTagline2, pdfCenter
    PDF.SetFont "Times", 10, False
    If Len(FirmAddressLine1) > 0 Then PDF.PrintText FirmAddressLine1, pdfCenter
    If Len(FirmAddressLine2) > 0 Then PDF.PrintText FirmAddressLine2, pdfCenter
    PDF.PrintText ""
    PDF.PrintText String$(RuleWidth, Chr$(175)), pdfCenter
    PDF.PrintText ""
    PDF.SetFont "Times", 12, False
End Sub

' --- Shared template piece #2: the name/firm/address signature block -------
' Used both for a letter's closing signature and for the Information
' Subpoena's own "who this is from" block, so both stay in sync with one
' another. IndentIn optionally indents the whole block (e.g. 4.5" to push it
' to the right of a "DATE:" line), matching SetStartX's units (inches).
Public Sub PrintFirmSignatureBlock(ByRef PDF As clsPDF, _
    Optional ByVal SignerName As String = "", _
    Optional ByVal SignerTitle As String = "", _
    Optional ByVal FirmName As String = FIRM_NAME, _
    Optional ByVal FirmTagline1 As String = "", _
    Optional ByVal FirmAddressLine1 As String = FIRM_ADDRESS_LINE1, _
    Optional ByVal FirmAddressLine2 As String = FIRM_ADDRESS_LINE2, _
    Optional ByVal FirmPhone As String = "", _
    Optional ByVal IndentIn As Double = 0)

    If Len(SignerName) > 0 Then
        If IndentIn > 0 Then PDF.SetStartX IndentIn
        PDF.PrintText SignerName, pdfLeft
    End If
    If Len(SignerTitle) > 0 Then
        If IndentIn > 0 Then PDF.SetStartX IndentIn
        PDF.PrintText SignerTitle, pdfLeft
    End If
    If IndentIn > 0 Then PDF.SetStartX IndentIn
    PDF.PrintText FirmName, pdfLeft
    If Len(FirmTagline1) > 0 Then
        If IndentIn > 0 Then PDF.SetStartX IndentIn
        PDF.PrintText FirmTagline1, pdfLeft
    End If
    If IndentIn > 0 Then PDF.SetStartX IndentIn
    PDF.PrintText FirmAddressLine1, pdfLeft
    If IndentIn > 0 Then PDF.SetStartX IndentIn
    PDF.PrintText FirmAddressLine2, pdfLeft
    If Len(FirmPhone) > 0 Then
        If IndentIn > 0 Then PDF.SetStartX IndentIn
        PDF.PrintText "PHONE: " & FirmPhone, pdfLeft
    End If
End Sub

' --- Shared template piece #4: the right-hand attorney/staff roster sidebar -
' Draws the divider line, firm name/tagline/address, and the attorney/staff
' roster down the right side of the CURRENT page (call it right after
' CreatePDF, before any NewPage, so it lands on page 1 - see PrintTextAt's
' notes in clsPDF.cls for why page placement works this way).
'
' The roster itself is fully data-driven from the employee table - it is NOT
' hardcoded here. Who appears, in what order, under what heading, and with
' what suffix (e.g. "Esq") is controlled by these columns on employee:
'   ON_LETTERHEAD, LH_ORDER, LH_HEADING, LH_SUFFIX, LH_SUFFIX_ITALIC,
'   LH_BLANK_BEFORE
' See DOC\Setup_LetterheadRoster.sql for column meanings and how to add/
' remove/reorder someone (it's just an UPDATE statement - no code changes,
' no PDF re-extraction, ever). That script also seeded the columns to match
' the roster in place the day this feature was built.
Public Sub PrintFirmSidebar(ByRef PDF As clsPDF)
    Dim rs As ADODB.Recordset
    Dim y As Double
    Dim boxW As Double
    Dim lastHeading As String
    Dim heading As String
    Dim fullName As String
    Dim lineText As String

    boxW = SIDEBAR_RIGHT - SIDEBAR_LEFT
    y = SIDEBAR_TOP

    PDF.DrawLine SIDEBAR_DIVIDER_X, SIDEBAR_TOP, SIDEBAR_DIVIDER_X, SIDEBAR_DIVIDER_BOTTOM

    PDF.SetFont "Times", 15, True
    PDF.PrintTextAt SIDEBAR_LEFT, y, FIRM_NAME, pdfCenter, boxW
    y = y + 0.32

    PDF.SetFont "Times", 9, True
    PDF.PrintTextAt SIDEBAR_LEFT, y, FIRM_TAGLINE1, pdfCenter, boxW
    y = y + 0.16

    PDF.SetFont "Times", 8, False, True
    PDF.PrintTextAt SIDEBAR_LEFT, y, FIRM_TAGLINE2, pdfCenter, boxW
    y = y + 0.22

    PDF.SetFont "Times", 8, False
    PDF.PrintTextAt SIDEBAR_LEFT, y, FIRM_ADDRESS_LINE1, pdfCenter, boxW
    y = y + 0.14
    PDF.PrintTextAt SIDEBAR_LEFT, y, FIRM_ADDRESS_LINE2, pdfCenter, boxW
    y = y + 0.14
    PDF.PrintTextAt SIDEBAR_LEFT, y, "Tel: " & FIRM_PHONE, pdfCenter, boxW
    y = y + 0.30

    Set rs = conn.Execute("SELECT FN, MI, LN, LH_HEADING, LH_SUFFIX, LH_SUFFIX_ITALIC, LH_BLANK_BEFORE " & _
        "FROM employee WHERE ON_LETTERHEAD = 'Y' ORDER BY LH_ORDER")

    lastHeading = "~~none~~"
    Do While Not rs.EOF
        If Trim$(rs!LH_BLANK_BEFORE & "") = "Y" Then y = y + 0.14

        heading = Trim$(rs!LH_HEADING & "")
        If heading <> "" And heading <> lastHeading Then
            PDF.SetFont "Times", 9, True, False, True
            PDF.PrintTextAt SIDEBAR_LEFT, y, heading, pdfCenter, boxW
            y = y + 0.16
        End If
        lastHeading = heading

        fullName = Trim$(rs!FN & "")
        If Trim$(rs!MI & "") <> "" Then fullName = fullName & " " & Trim$(rs!MI & "") & "."
        fullName = fullName & " " & Trim$(rs!LN & "")

        lineText = fullName
        If Trim$(rs!LH_SUFFIX & "") <> "" Then
            If Trim$(rs!LH_SUFFIX_ITALIC & "") = "Y" Then
                lineText = lineText & ", *" & Trim$(rs!LH_SUFFIX & "") & "*"
            Else
                lineText = lineText & ", " & Trim$(rs!LH_SUFFIX & "")
            End If
        End If

        PDF.SetFont "Times", 9, False
        PDF.PrintTextAt SIDEBAR_LEFT, y, lineText, pdfCenter, boxW
        y = y + 0.155

        rs.MoveNext
    Loop
    rs.Close

    PDF.SetFont "Times", 12, False   ' restore the default before the caller resumes body text
End Sub

' --- Shared template piece #3: the bottom debt-collector disclosure --------
Public Function LetterheadFooterText(Optional ByVal FirmName As String = FIRM_NAME) As String
    LetterheadFooterText = "THE LAW FIRM OF " & UCase$(FirmName) & " IS A DEBT COLLECTOR " & _
        "ATTEMPTING TO COLLECT A DEBT AND ANY INFORMATION OBTAINED " & _
        "WILL BE USED FOR THAT PURPOSE."
End Function

Public Sub PrintFDCPADisclosure(ByRef PDF As clsPDF, Optional ByVal FirmName As String = FIRM_NAME)
    PDF.SetFont "Times", 12, True
    PDF.PrintTextCenteredOnFullPage LetterheadFooterText(FirmName), pdfCenter
    PDF.SetFont "Times", 12, False
End Sub

' Prints a letter on the firm's letterhead. Only RecipientLines, DateStr,
' Salutation, Body, SignerName are required - everything else has a sensible
' default so you can call this the same way every time and just change the
' body.
Public Sub PrintLetterhead( _
    ByVal PdfPath As String, _
    ByVal RecipientLines As String, _
    ByVal DateStr As String, _
    ByVal Salutation As String, _
    ByVal Body As String, _
    ByVal SignerName As String, _
    Optional ByVal ReferenceLine As String = "", _
    Optional ByVal SignerTitle As String = "", _
    Optional ByVal ClosingPhrase As String = "Very truly yours,", _
    Optional ByVal Enclosures As Boolean = False, _
    Optional ByVal FooterDisclosure As String = "", _
    Optional ByVal FirmName As String = FIRM_NAME, _
    Optional ByVal FirmTagline1 As String = FIRM_TAGLINE1, _
    Optional ByVal FirmTagline2 As String = FIRM_TAGLINE2, _
    Optional ByVal FirmAddressLine1 As String = FIRM_ADDRESS_LINE1, _
    Optional ByVal FirmAddressLine2 As String = FIRM_ADDRESS_LINE2)

    Dim PDF As New clsPDF
    Dim recipLine As Variant
    Dim recipArr() As String
    Dim paraArr() As String
    Dim I As Integer

    PDF.CreatePDF PdfPath
    PDF.SetMargins 1, 0.75, 1, 0.75

    ' --- Letterhead header (shared template piece - see top of module) ------
    PrintFirmHeader PDF, FirmName, FirmTagline1, FirmTagline2, FirmAddressLine1, FirmAddressLine2

    ' --- Date / recipient / reference / salutation ---------------------------
    PDF.PrintText DateStr, pdfLeft
    PDF.PrintText ""

    recipArr = Split(RecipientLines, vbNewLine)
    For I = 0 To UBound(recipArr)
        If Len(Trim$(recipArr(I))) > 0 Then PDF.PrintText Trim$(recipArr(I)), pdfLeft
    Next I
    PDF.PrintText ""

    If Len(ReferenceLine) > 0 Then
        PDF.PrintText "Re:  " & ReferenceLine, pdfLeft
        PDF.PrintText ""
    End If

    PDF.PrintText Salutation, pdfLeft
    PDF.PrintText ""

    ' --- Body: split on blank lines into paragraphs, justify each one --------
    paraArr = Split(Replace$(Body, vbCrLf, vbNewLine), vbNewLine & vbNewLine)
    For I = 0 To UBound(paraArr)
        If Len(Trim$(paraArr(I))) > 0 Then
            PDF.PrintParagraph Trim$(paraArr(I)), pdfJustify
            PDF.PrintText ""
        End If
    Next I

    ' --- Closing / signature ---------------------------------------------------
    ' (Just the person's name/title - the firm identity is already shown once
    ' at the top via PrintFirmHeader, no need to repeat the full firm block
    ' here. Documents that DO want the full name+firm+address block near a
    ' signature, like the Information Subpoena, use PrintFirmSignatureBlock
    ' directly - see LET10_InfoSubpoena_clsPDF.bas.)
    PDF.PrintText ClosingPhrase, pdfLeft
    PDF.PrintText ""
    PDF.PrintText ""
    PDF.PrintText SignerName, pdfLeft
    If Len(SignerTitle) > 0 Then PDF.PrintText SignerTitle, pdfLeft
    If Enclosures Then PDF.PrintText "Enclosure(s)", pdfLeft

    ' --- Optional bottom disclosure (e.g. FDCPA debt-collector notice) -------
    If Len(FooterDisclosure) > 0 Then
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.SetFont "Times", 10, True
        PDF.PrintParagraph FooterDisclosure, pdfCenter
    End If

    PDF.SavePDF
End Sub

' Demo call showing the intended everyday usage: change only the body (and
' recipient/date/salutation/signer as needed) each time you need a new
' letter - the letterhead itself never changes.
Public Sub Demo_PrintLetterhead()
    PrintLetterhead _
        PdfPath:="C:\temp\SampleLetter.pdf", _
        RecipientLines:="John Doe" & vbNewLine & "123 Main St." & vbNewLine & "Anytown, NY 12345", _
        DateStr:=Format$(Date, "mmmm d, yyyy"), _
        Salutation:="Dear Mr. Doe:", _
        ReferenceLine:="Account No. 12345-6789", _
        Body:= _
            "This is the first paragraph of the letter body. Everything above this line " & _
            "(the firm name, tagline, address, date, recipient block, reference line, and " & _
            "salutation) is handled automatically by PrintLetterhead - you only need to " & _
            "supply this paragraph text." & vbNewLine & vbNewLine & _
            "This is a second paragraph, to show that multiple paragraphs are separated " & _
            "automatically by leaving a blank line between them in the Body string, and " & _
            "that each one word-wraps and justifies on its own.", _
        SignerName:="Amanda C Shaw", _
        SignerTitle:="Attorney", _
        Enclosures:=True

    MsgBox "Saved C:\temp\SampleLetter.pdf"
End Sub

' Opens a PDF for viewing without Adobe "creating file to read" / save-on-close
' prompts. Uses Microsoft Edge or Acrobat Reader only — never the Windows default
' handler (Adobe Acrobat Pro), which prompts to save on close.
Public Sub OpenPdfForViewing(ByVal pdfPath As String)
    Dim edgePath As String
    Dim readerPath As String
    Dim wsh As Object

    pdfPath = Trim$(pdfPath)
    If Len(pdfPath) = 0 Then Exit Sub
    If Len(Dir$(pdfPath)) = 0 Then Exit Sub

    Set wsh = CreateObject("WScript.Shell")

    edgePath = ResolveEdgeExe()
    If Len(edgePath) > 0 Then
        wsh.Run """" & edgePath & """ --new-window --no-first-run """ & pdfPath & """", 1, False
        Exit Sub
    End If

    readerPath = FindAcrobatReaderExe()
    If Len(readerPath) > 0 Then
        wsh.Run """" & readerPath & """ /s /n """ & pdfPath & """", 1, False
        Exit Sub
    End If

    MsgBox "PDF saved (not opened automatically):" & vbCrLf & pdfPath & vbCrLf & vbCrLf & _
        "Edge or Adobe Reader was not found. Open the file manually if needed.", vbInformation, "PDF"
End Sub

Private Function ResolveEdgeExe() As String
    Dim p As String

    p = Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe"
    If Len(Dir$(p)) > 0 Then
        ResolveEdgeExe = p
        Exit Function
    End If

    p = Environ("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe"
    If Len(Dir$(p)) > 0 Then
        ResolveEdgeExe = p
        Exit Function
    End If

    p = Environ("LOCALAPPDATA") & "\Microsoft\Edge\Application\msedge.exe"
    If Len(Dir$(p)) > 0 Then ResolveEdgeExe = p
End Function

Private Function FindAcrobatReaderExe() As String
    Dim candidates(0 To 3) As String
    Dim I As Integer

    candidates(0) = Environ("ProgramFiles(x86)") & "\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    candidates(1) = Environ("ProgramFiles") & "\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    candidates(2) = Environ("ProgramFiles(x86)") & "\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    candidates(3) = Environ("ProgramFiles") & "\Adobe\Acrobat Reader\AcroRd32.exe"

    For I = 0 To 3
        If Len(Dir$(candidates(I))) > 0 Then
            FindAcrobatReaderExe = candidates(I)
            Exit Function
        End If
    Next I
End Function
