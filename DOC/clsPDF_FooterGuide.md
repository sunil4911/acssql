# clsPDF Footer Guide

This guide explains how to add the bottom disclosure/footer to PDF documents
created with `clsPDF`. The footer is **real text** (Times 12 bold, centered) —
not vector artwork — so you can change the wording per document without
re-extracting `LH_Content.txt` or editing `LH.pdf`.

There are **two ways** to put footer text on a PDF. Pick the one that matches
how you want the footer to behave:

| Method | Where it appears | Needs letterhead? | Best for |
|---|---|---|---|
| `SetLetterheadFooter` | Fixed position near the bottom of **page 1** | No | Letters, invoices, anything that should look like pre-printed letterhead paper |
| `PrintFDCPADisclosure` | Flows after your last line of content | No | Reports, subpoena packages, multi-page docs where the footer follows the body |

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [SetLetterheadFooter — Fixed Bottom Footer](#setletterheadfooter--fixed-bottom-footer)
3. [Changing the Footer Text](#changing-the-footer-text)
4. [Using the Footer Without Letterhead](#using-the-footer-without-letterhead)
5. [Using the Footer With Letterhead](#using-the-footer-with-letterhead)
6. [Adjusting Vertical Position](#adjusting-vertical-position)
7. [Turning the Footer Off](#turning-the-footer-off)
8. [PrintFDCPADisclosure — Flowing Footer](#printfdcpadisclosure--flowing-footer)
9. [LetterheadFooterText Helper](#letterheadfootertext-helper)
10. [Common Recipes](#common-recipes)
11. [API Reference](#api-reference)

---

## Quick Start

**Fixed footer on page 1 (most common for letters):**

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Letter.pdf"
PDF.SetLetterheadFooter LetterheadFooterText()
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Dear John,"
' ... body ...
PDF.SavePDF
```

**Flowing footer after your content:**

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Report.pdf"
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Report body here..."
PrintFDCPADisclosure PDF
PDF.SavePDF
```

Both approaches work **with or without** `SetLetterhead`.

---

## SetLetterheadFooter — Fixed Bottom Footer

```vb
PDF.SetLetterheadFooter FooterText, [YFromTopIn]
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `FooterText` | `String` | *(required)* | The full footer wording. Empty string = no footer. |
| `YFromTopIn` | `Double` | `10.35` | Vertical position of the **first line**, in inches from the top of the page |

**How it works:**

1. You call `SetLetterheadFooter` any time after `CreatePDF` and before `SavePDF`.
2. The text is held in memory until `SavePDF`.
3. At save time, the footer is drawn onto **page 1 only**, centered between the
   left and right margins, with automatic word-wrapping.
4. It does **not** move the normal text cursor — your body text flows
   independently above it.

**Important:** Call `SetMargins` before `SavePDF`. The footer centers itself
within the margin width. If margins are still at the default (1"), that works
fine; just set them explicitly if you need a custom layout.

---

## Changing the Footer Text

### Option A — Standard FDCPA wording (default firm name)

Uses the `FIRM_NAME` constant from `Letterhead_clsPDF.bas`:

```vb
PDF.SetLetterheadFooter LetterheadFooterText()
```

Produces:

> THE LAW FIRM OF RIEHLMAN, SHAFER & SHAW, LLC IS A DEBT COLLECTOR ATTEMPTING TO COLLECT A DEBT AND ANY INFORMATION OBTAINED WILL BE USED FOR THAT PURPOSE.

### Option B — Standard wording, different firm name

```vb
PDF.SetLetterheadFooter LetterheadFooterText("Smith & Jones, LLC")
```

Only the firm name in the sentence changes; the rest of the wording stays the same.

### Option C — Fully custom text

```vb
PDF.SetLetterheadFooter "YOUR ENTIRE CUSTOM FOOTER WORDING GOES HERE."
```

Use this when the disclosure wording is completely different from the default.

### Option D — Change the default for everyone

Edit **one place** in `Letterhead_clsPDF.bas`:

```vb
Public Const FIRM_NAME As String = "Riehlman, Shafer & Shaw, LLC"
```

Or edit the `LetterheadFooterText` function itself if you want different
standard sentence structure project-wide.

---

## Using the Footer Without Letterhead

The footer is fully independent of the letterhead sidebar. Any document can
have just the footer:

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Report.pdf"
PDF.SetLetterheadFooter LetterheadFooterText()
PDF.SetMargins 1, 1, 1, 1

PDF.SetFont "Times", 12
PDF.PrintText "Annual Summary Report", pdfCenter
PDF.PrintText ""
PDF.PrintParagraph "This report contains account activity for the period...", pdfJustify

PDF.SavePDF
```

No call to `SetLetterhead` is required. The sidebar graphic is simply omitted;
only the bottom disclosure appears on page 1.

---

## Using the Footer With Letterhead

For firm letters that use the pre-extracted sidebar artwork from `LH_Content.txt`:

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Letter.pdf"
PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"       ' sidebar + divider (static graphic)
PDF.SetLetterheadFooter LetterheadFooterText()          ' bottom disclosure (variable text)
PDF.SetMargins 1, 1, 1, 1

PDF.SetFont "Times", 12
PDF.PrintText Format$(Date, "mmmm d, yyyy"), pdfLeft
PDF.PrintText ""
PDF.PrintText "Dear Mr. Smith:", pdfLeft
PDF.PrintParagraph "Body of the letter...", pdfJustify

PDF.SavePDF
```

This is exactly what `Command8_Click` in `ACSSQLMAIN.frm` demonstrates.

The sidebar names are still baked into `LH_Content.txt` (vector artwork). Only
the **bottom** text is variable. If you also need variable sidebar names, see
`PrintFirmSidebar` in `Letterhead_clsPDF.bas` and `DOC\Setup_LetterheadRoster.sql`.

---

## Adjusting Vertical Position

By default the first footer line starts at **10.35 inches** from the top of the
page (near the bottom of an 11" letter). Move it up or down with the optional
second parameter:

```vb
PDF.SetLetterheadFooter LetterheadFooterText(), 10.0   ' slightly higher
PDF.SetLetterheadFooter LetterheadFooterText(), 10.5   ' slightly lower
```

If the footer wraps to two lines, subsequent lines stack downward from that
starting position. Adjust `YFromTopIn` if the footer overlaps your body text or
sits too close to the page edge.

---

## Turning the Footer Off

**For one document** — pass an empty string:

```vb
PDF.SetLetterheadFooter ""
```

**Explicitly clear** a footer set earlier on the same `clsPDF` object:

```vb
PDF.ClearLetterheadFooter
```

**Letterhead without footer** — use `SetLetterhead` only, skip `SetLetterheadFooter`:

```vb
PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
' no SetLetterheadFooter call
```

---

## PrintFDCPADisclosure — Flowing Footer

When you want the disclosure to appear **after your content** (not pinned to
the bottom of page 1), use the helper in `Letterhead_clsPDF.bas`:

```vb
PrintFDCPADisclosure PDF                          ' default firm name
PrintFDCPADisclosure PDF, "Other Firm Name, LLC"  ' custom firm name
```

**Differences from `SetLetterheadFooter`:**

| | `SetLetterheadFooter` | `PrintFDCPADisclosure` |
|---|---|---|
| Position | Fixed near bottom of page 1 | After the last line you printed |
| Page | Page 1 only | Whatever page the cursor is on |
| When to call | Any time before `SavePDF` | Right before `SavePDF`, after body |
| Affects text cursor | No | Yes (advances cursor) |

**Example — subpoena package (page 2 footer):**

```vb
PDF.CreatePDF "C:\temp\Package.pdf"
PDF.SetMargins 1, 0.75, 1, 0.75

' Page 1: cover letter body ...
PDF.NewPage

' Page 2: subpoena content ...
PrintFDCPADisclosure PDF    ' footer lands at end of page 2 content

PDF.SavePDF
```

---

## LetterheadFooterText Helper

Defined in `Letterhead_clsPDF.bas`:

```vb
Public Function LetterheadFooterText(Optional ByVal FirmName As String = FIRM_NAME) As String
```

Returns the standard FDCPA disclosure string with the firm name in uppercase.
Use it anywhere you need the default wording:

```vb
' In SetLetterheadFooter:
PDF.SetLetterheadFooter LetterheadFooterText()

' Or build your own string from it:
Dim s As String
s = LetterheadFooterText()
s = Replace(s, "DEBT COLLECTOR", "DEBT COLLECTOR AND CREDITOR")
PDF.SetLetterheadFooter s
```

---

## Common Recipes

### Recipe 1 — Simple letter with letterhead + footer

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Letter.pdf"
PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
PDF.SetLetterheadFooter LetterheadFooterText()
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Dear Client:", pdfLeft
PDF.PrintParagraph "Your letter body here.", pdfJustify
PDF.SavePDF
```

### Recipe 2 — Invoice with footer, no letterhead sidebar

```vb
Dim PDF As New clsPDF
PDF.CreatePDF "C:\temp\Invoice.pdf"
PDF.SetLetterheadFooter LetterheadFooterText()
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 16, True
PDF.PrintText "Invoice", pdfCenter
' ... line items ...
PDF.SavePDF
```

### Recipe 3 — Multi-page document, footer only on last page

Do **not** use `SetLetterheadFooter` (it is always page 1). Use the flowing
approach instead:

```vb
PDF.CreatePDF "C:\temp\LongDoc.pdf"
PDF.SetMargins 1, 1, 1, 1
' ... pages of content ...
PDF.NewPage
' ... more content ...
PrintFDCPADisclosure PDF
PDF.SavePDF
```

### Recipe 4 — Reuse one clsPDF object for many letters

```vb
Dim PDF As New clsPDF
PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
PDF.SetLetterheadFooter LetterheadFooterText()

' Letter 1
PDF.CreatePDF "C:\temp\Letter1.pdf"
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Dear Alice,", pdfLeft
PDF.SavePDF

' Letter 2 — same letterhead + footer, different body
PDF.CreatePDF "C:\temp\Letter2.pdf"
PDF.SetMargins 1, 1, 1, 1
PDF.SetFont "Times", 12
PDF.PrintText "Dear Bob,", pdfLeft
PDF.SavePDF
```

`SetLetterhead` and `SetLetterheadFooter` stay in memory until you call
`ClearLetterhead` or `ClearLetterheadFooter`.

### Recipe 5 — Different footer per document

```vb
PDF.CreatePDF "C:\temp\DocA.pdf"
PDF.SetLetterheadFooter LetterheadFooterText("Firm A, LLC")
' ... content ...
PDF.SavePDF

PDF.CreatePDF "C:\temp\DocB.pdf"
PDF.SetLetterheadFooter LetterheadFooterText("Firm B, LLC")
' ... content ...
PDF.SavePDF
```

---

## API Reference

### clsPDF.cls

| Method | Signature | Description |
|---|---|---|
| `SetLetterheadFooter` | `SetLetterheadFooter(FooterText As String, [YFromTopIn = 10.35])` | Stores footer text; drawn on page 1 at `SavePDF`. |
| `ClearLetterheadFooter` | `ClearLetterheadFooter()` | Disables the footer for subsequent saves on this object. |

### Letterhead_clsPDF.bas

| Method | Signature | Description |
|---|---|---|
| `LetterheadFooterText` | `LetterheadFooterText([FirmName = FIRM_NAME]) As String` | Returns the standard FDCPA disclosure string. |
| `PrintFDCPADisclosure` | `PrintFDCPADisclosure(PDF, [FirmName = FIRM_NAME])` | Prints the disclosure in the normal text flow (centered, justified wrap). |

### Related (letterhead sidebar only)

| Method | Description |
|---|---|
| `SetLetterhead` | Loads the static sidebar/divider graphic from `LH_Content.txt`. |
| `ClearLetterhead` | Turns off the sidebar graphic. |
| `PrintFirmSidebar` | Database-driven sidebar (variable names from `employee` table). |

---

## See Also

- `clsPDF_UserGuide.md` — full clsPDF API (margins, fonts, tabs, letterhead, etc.)
- `Letterhead_clsPDF.bas` — shared letterhead template module
- `DOC\Setup_LetterheadRoster.sql` — database-driven sidebar roster setup
- `DOC\Extract_LH_Content.ps1` — regenerates `LH_Content.txt` when `LH.pdf` changes
