# clsPDF User Guide

`clsPDF` (`clsPDF.cls`) is a hand-rolled PDF writer for VB6 - it needs no external
libraries, printer drivers, or ActiveX controls. You create one `clsPDF` object,
call methods on it to lay out text, and then save a real `.pdf` file to disk.

This guide covers every public method with runnable examples. For the full
list of private/internal helpers, see the comments inside `clsPDF.cls` itself.

## Table of Contents

1. [Quick Start](#quick-start)
2. [Creating the PDF and Setting Margins](#creating-the-pdf-and-setting-margins)
3. [Fonts](#fonts)
4. [Inline Text Styling (bold / italic / underline)](#inline-text-styling)
5. [Paragraphs and Alignment](#paragraphs-and-alignment)
6. [Columns and Tables](#columns-and-tables)
7. [One-off Horizontal Placement (signature lines, etc.)](#one-off-horizontal-placement)
8. [Page Breaks](#page-breaks)
9. [Letterhead Background](#letterhead-background)
10. [Footer / Bottom Disclosure](clsPDF_FooterGuide.md) — fixed and flowing options
11. [Full Worked Example: an Invoice](#full-worked-example-an-invoice)
12. [Troubleshooting](#troubleshooting)
13. [Public API Reference](#public-api-reference)

---

## Quick Start

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\TEST.PDF"
PDF.SetMargins 1, 1, 1, 1           ' inches: left, top, right, bottom
PDF.SetFont "Times", 12, False      ' family, size, bold, [italic]
PDF.PrintText "Dear John,", pdfLeft
PDF.PrintText ""                    ' blank line
PDF.PrintText "Please pay **$500.00** by Friday.", pdfJustify
PDF.NewPage                         ' force a page break
PDF.PrintText "Sincerely,", pdfRight
PDF.SavePDF
```

Every `clsPDF` usage follows the same shape:

1. `New clsPDF`
2. `CreatePDF` (once)
3. Any mix of `SetMargins`, `SetFont`, `PrintText`/`PrintParagraph`,
   `PrintTabbedLine`, `NewPage`, etc.
4. `SavePDF` (once, at the end)

---

## Creating the PDF and Setting Margins

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Letter.pdf"   ' starts a new US-Letter (8.5"x11") page

PDF.SetMargins 1, 1, 1, 1            ' left, top, right, bottom - in inches
' or, if you prefer points (1/72 inch):
PDF.SetMarginsPt 72, 72, 72, 72      ' same as above, in points
```

`CreatePDF` must be called before any other method (`SetFont`, `PrintText`,
etc.) - calling them first raises a runtime error telling you so.

---

## Fonts

```vb
PDF.SetFont "Times", 12, False              ' family, size, bold
PDF.SetFont "Helvetica", 10, True, True      ' family, size, bold, italic
PDF.SetFont "Courier", 11, False, False, True ' family, size, bold, italic, underline
```

Recognized family names (case-insensitive): `Times`/`Times New Roman`/
`Times-Roman`/`serif`, `Helvetica`/`Arial`/`sans`/`sans-serif`, `Courier`/
`Courier New`/`mono`/`monospace`. Anything unrecognized falls back to
Helvetica.

You can also use direct PDF font-resource names, if you prefer that style:

```vb
PDF.SetFont "F1", 12   ' F1 = Helvetica, F2 = Helvetica-Bold, ... F12 = Courier-BoldOblique
```

`SetFont` only affects text printed *after* it's called - it doesn't
retroactively restyle anything already printed.

---

## Inline Text Styling

`PrintText`, `PrintParagraph`, and `PrintTabbedLine` all understand simple
inline markup inside the string you pass them, similar to Markdown:

| Markup | Effect |
|---|---|
| `**bold**` | bold |
| `*italic*` | italic |
| `__underline__` | underline (drawn as an actual line under the text) |

They combine freely:

```vb
PDF.PrintText "This is ***bold italic***, and this is **__bold underline__**."
```

Each marker *toggles* relative to whatever `SetFont`'s `bold`/`italic`/
`underline` state currently is. If a literal `*` or `_` needs to appear in
your text (e.g. `"3*4"`), it will be misread as markup - use `SetFont`'s
boolean parameters instead of markup in that situation.

---

## Paragraphs and Alignment

```vb
PDF.PrintText "Some long paragraph text that will automatically word-wrap " & _
              "to fit inside the margins...", pdfJustify
```

`PrintParagraph` is an identical alias for `PrintText` - use whichever name
reads better in your code.

Alignment options (the `PdfAlign` enum):

| Constant | Effect |
|---|---|
| `pdfLeft` | left-aligned |
| `pdfCenter` | centered |
| `pdfRight` | right-aligned |
| `pdfJustify` | *(default)* stretches word spacing to reach the right margin; the last wrapped line of a paragraph is left-aligned, per normal typographic convention |

```vb
PDF.PrintText "Centered Title", pdfCenter
PDF.PrintText ""                       ' blank line - just leaves vertical space
```

---

## Columns and Tables

Use `SetTabStops` (inches) or `SetTabStopsPt` (points) to define column
positions, then `PrintTabbedLine` to print a row - split the row's text on
`vbTab`, once per column:

```vb
PDF.SetTabStops 0, 2.5, 4.5            ' 3 columns: 0in, 2.5in, 4.5in from the left margin
PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price"
PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**"
```

By default every column in a `PrintTabbedLine` call is left-aligned. For a
typical table where the label column stays left-aligned but numeric columns
are right-aligned (so numbers/currency line up nicely), pass an `Align`
**per column** using `Array(...)` instead of a single value - and use the
*same* array on both the header row and the data rows so the headings line
up directly above their column:

```vb
PDF.SetTabStops 0, 2.5, 4.5
PDF.SetFont "Helvetica", 11, True
PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price", _
    Array(pdfLeft, pdfRight, pdfRight)

PDF.SetFont "Helvetica", 11, False
PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**", _
    Array(pdfLeft, pdfRight, pdfRight)
PDF.PrintTabbedLine "Gadget" & vbTab & "1" & vbTab & "**$14.99**", _
    Array(pdfLeft, pdfRight, pdfRight)
```

Notes:

- You can still pass a single alignment (e.g. `PrintTabbedLine text, pdfRight`)
  to apply it to every column in that row - this is unchanged, older behavior.
- If the `Array(...)` has fewer entries than columns, the last entry repeats
  for any remaining columns.
- You can use `vbTab` as many times as you like on one line to create more
  than 3 columns - just make sure `SetTabStops` defines a position for each
  one (extra columns beyond the defined stops just flow left-to-right with a
  small gap instead of landing at a specific position):

```vb
PDF.SetTabStops 0, 2.5, 4, 5.5
PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Unit" & vbTab & "Price", _
    Array(pdfLeft, pdfRight, pdfRight, pdfRight)
```

- Column segments support the same `**bold**`/`*italic*`/`__underline__`
  markup as `PrintText`, but are **not** word-wrapped - keep each cell short.

---

## One-off Horizontal Placement

For something like a signature line mid-page, without setting up tab stops:

```vb
PDF.SetStartX 3.5                 ' inches from the left margin
PDF.PrintText "X_______________"
' (flow automatically reverts to the normal left margin after this call)
```

`SetStartXPt` is the points-based equivalent. It only affects the very next
`PrintText`/`PrintParagraph` call's first line - if that text wraps to
additional lines, the wrapped lines return to the normal left margin, and
every call after that is unaffected.

---

## Page Breaks

```vb
PDF.NewPage
```

Forces a page break immediately. `clsPDF` also inserts page breaks
automatically whenever the next line would run past the bottom margin - you
don't need to call `NewPage` yourself except when you deliberately want
content to start on a fresh page (e.g. between sections of a long report).

---

## Letterhead Background

`SetLetterhead` stamps a pre-designed graphic - a firm's letterhead artwork,
e.g. a logo/heading, a divider line, a footer disclosure, etc. - onto the
**first page only** of the next document(s) saved by this `clsPDF` instance,
the way pre-printed letterhead paper works:

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Letter.pdf"
PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
PDF.SetLetterheadFooter LetterheadFooterText()   ' variable bottom disclosure
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Dear John,"
' ... rest of the letter ...
PDF.SavePDF
```

`ClearLetterhead` turns off the sidebar graphic; `ClearLetterheadFooter` turns off
the bottom text. Since both are held in memory after you set them once, you can
reuse the same `clsPDF` object across many letters without re-reading files.

### How it works, and its one limitation

`SetLetterhead` does **not** take a PDF path - it takes a path to a small
plain-text file of raw PDF drawing commands, pre-extracted from a source PDF
(such as `DOC\LH.pdf`) by the companion script `DOC\Extract_LH_Content.ps1`.
This is because VB6 has no built-in way to decompress a PDF's internal
`FlateDecode`-compressed content streams, so that decompression - along with
the geometric analysis needed to keep only the letterhead artwork and discard
whatever one-off sample letter happened to be on that source page - is done
once, ahead of time, by that PowerShell script instead of at runtime.

This means the technique only works when the source PDF's page is pure
vector artwork (paths/fills - typically what you get when a "Print to PDF"
driver converts everything, including text, to outline shapes) with an empty
`/Resources` dictionary (no embedded fonts or images). `DOC\LH.pdf` satisfies
this. If you ever need to swap in a different letterhead design, replace
`DOC\LH.pdf` and re-run:

```powershell
powershell -File C:\ACSSQL\DOC\Extract_LH_Content.ps1
```

That regenerates `DOC\LH_Content.txt` (the file `SetLetterhead` reads). The
script keeps only the **sidebar and divider line** by default; the bottom
FDCPA disclosure is **not** baked into that file anymore — supply it per
document with `SetLetterheadFooter`, or use the ready-made
`LetterheadFooterText()` helper in `Letterhead_clsPDF.bas`:

```vb
PDF.SetLetterheadFooter LetterheadFooterText()              ' default firm name
PDF.SetLetterheadFooter LetterheadFooterText("Other Firm")  ' different firm
PDF.SetLetterheadFooter "Your own custom footer wording."    ' fully custom
```

If the new artwork's "keep" regions differ from the current sidebar-on-the-right
layout, adjust the `-SidebarMinX` parameter at the top of that script (see the
comments inside it for how that is used).

For full documentation on footer options (with or without letterhead, custom
text, vertical position, flowing vs. fixed), see **`DOC\clsPDF_FooterGuide.md`**.

---

## Full Worked Example: an Invoice

This is the same example used in `ACSSQLMAIN.frm` (`Command8_Click`):

```vb
Private Sub Command8_Click()
    Dim PDF As New clsPDF
    PDF.CreatePDF "C:\temp\Invoice.pdf"
    PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
    PDF.SetLetterheadFooter LetterheadFooterText()
    PDF.SetMargins 1, 1, 1, 1

    PDF.SetFont "Times", 16, True
    PDF.PrintText "Invoice", pdfCenter

    PDF.SetFont "Helvetica", 12
    PDF.PrintText ""
    PDF.PrintText "Dear John, please review the *itemized* charges below.", pdfJustify
    PDF.PrintText ""

    PDF.SetTabStops 0, 3, 4.5
    PDF.SetFont "Helvetica", 11, True
    PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price", _
        Array(pdfLeft, pdfRight, pdfRight)
    PDF.SetFont "Helvetica", 11, False
    PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**", _
        Array(pdfLeft, pdfRight, pdfRight)
    PDF.PrintTabbedLine "Gadget" & vbTab & "1" & vbTab & "**$14.99**", _
        Array(pdfLeft, pdfRight, pdfRight)

    PDF.PrintText ""
    PDF.PrintText "Please remit payment within __10 days__ of this notice.", pdfJustify

    PDF.SavePDF
End Sub
```

This produces a one-page PDF with the firm's letterhead in the background, a
centered bold title, a justified intro paragraph with an italic word, a
3-column table with right-aligned numeric columns (headers lined up directly
above the data), and a closing line with an underlined phrase.

---

## Troubleshooting

**Runtime error breaking on `Open m_FilePath For Binary Access Write As #fnum`
inside `SavePDF`:** the output file is almost always still open in a PDF
viewer (Acrobat, browser tab, etc.) or another program. Close it there and
run `SavePDF` again. `SavePDF` now raises a clear message identifying the
file path when this happens, instead of a raw runtime error.

**`Err.Raise` saying "CreatePDF must be called before ...":** you called
`SetFont`/`PrintText`/etc. before `CreatePDF`. Always call `CreatePDF` first.

**Table columns don't line up:** make sure the header row and data rows pass
the *same* `SetTabStops` values and the *same* per-column `Array(...)`
alignment - see [Columns and Tables](#columns-and-tables).

---

## Public API Reference

| Method | Signature | Purpose |
|---|---|---|
| `CreatePDF` | `CreatePDF(sFilePath As String)` | Starts a new PDF at the given path. Must be called first. |
| `SetMargins` | `SetMargins(LeftIn, TopIn, RightIn, BottomIn As Double)` | Page margins, in inches. |
| `SetMarginsPt` | `SetMarginsPt(LeftPt, TopPt, RightPt, BottomPt As Double)` | Page margins, in points. |
| `SetTabStops` | `SetTabStops(Positions...)` | Column start positions for `PrintTabbedLine`, in inches. |
| `SetTabStopsPt` | `SetTabStopsPt(Positions...)` | Same, in points. |
| `SetStartX` | `SetStartX(XIn As Double)` | One-off left-start position (inches) for the next `PrintText`/`PrintParagraph` call. |
| `SetStartXPt` | `SetStartXPt(XPt As Double)` | Same, in points. |
| `NewPage` | `NewPage()` | Forces a page break. |
| `SetLetterhead` | `SetLetterhead(ContentFilePath As String)` | Stamps the letterhead graphic in `ContentFilePath` (see `DOC\LH_Content.txt`) onto page 1 of the next document(s). |
| `ClearLetterhead` | `ClearLetterhead()` | Turns off the letterhead set by `SetLetterhead`. |
| `SetLetterheadFooter` | `SetLetterheadFooter(FooterText As String, [YFromTopIn = 10.35])` | Draws variable bottom disclosure text on page 1 (inches from top for first line). |
| `ClearLetterheadFooter` | `ClearLetterheadFooter()` | Turns off the footer set by `SetLetterheadFooter`. |
| `SetFont` | `SetFont(FontName As String, Size As Double, [bold], [italic], [underline])` | Sets the current font. |
| `PrintText` | `PrintText(Text As String, [Align = pdfJustify])` | Prints a word-wrapped paragraph with inline markup support. |
| `PrintParagraph` | `PrintParagraph(Text As String, [Align = pdfJustify])` | Alias for `PrintText`. |
| `PrintTabbedLine` | `PrintTabbedLine(Text As String, [Align = pdfLeft])` | Prints one row of columns, split on `vbTab`. `Align` may be a single `PdfAlign` or `Array(...)` for per-column alignment. |
| `SavePDF` | `SavePDF()` | Finalizes and writes the PDF to disk. Must be called last. |

**`PdfAlign` enum:** `pdfLeft = 0`, `pdfCenter = 1`, `pdfRight = 2`,
`pdfJustify = 3`.
