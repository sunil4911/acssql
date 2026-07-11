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
9. [Full Worked Example: an Invoice](#full-worked-example-an-invoice)
10. [Troubleshooting](#troubleshooting)
11. [Public API Reference](#public-api-reference)

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

## Full Worked Example: an Invoice

This is the same example used in `ACSSQLMAIN.frm` (`Command8_Click`):

```vb
Private Sub Command8_Click()
    Dim PDF As New clsPDF
    PDF.CreatePDF "C:\temp\Invoice.pdf"
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

This produces a one-page PDF with a centered bold title, a justified intro
paragraph with an italic word, a 3-column table with right-aligned numeric
columns (headers lined up directly above the data), and a closing line with
an underlined phrase.

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
| `SetFont` | `SetFont(FontName As String, Size As Double, [bold], [italic], [underline])` | Sets the current font. |
| `PrintText` | `PrintText(Text As String, [Align = pdfJustify])` | Prints a word-wrapped paragraph with inline markup support. |
| `PrintParagraph` | `PrintParagraph(Text As String, [Align = pdfJustify])` | Alias for `PrintText`. |
| `PrintTabbedLine` | `PrintTabbedLine(Text As String, [Align = pdfLeft])` | Prints one row of columns, split on `vbTab`. `Align` may be a single `PdfAlign` or `Array(...)` for per-column alignment. |
| `SavePDF` | `SavePDF()` | Finalizes and writes the PDF to disk. Must be called last. |

**`PdfAlign` enum:** `pdfLeft = 0`, `pdfCenter = 1`, `pdfRight = 2`,
`pdfJustify = 3`.
