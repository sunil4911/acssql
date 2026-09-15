VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

  'MultiUse = -1  'True
  'Persistable = 0  'NotPersistable
  'DataBindingBehavior = 0  'vbNone
  'DataSourceBehavior = 0   'vbNone
  'MTSTransMode = 0  'NotAnMTSObject
'End
Option Explicit

' clsPDF - hand-rolled PDF writer (no external libraries/printer drivers).
'
' Usage:
'   Dim PDF As clsPDF
'   Set PDF = New clsPDF
'   PDF.CreatePDF "C:\temp\TEST.PDF"
'   PDF.SetMargins 1, 1, 1, 1          ' inches: left, top, right, bottom
'   PDF.SetFont "Times", 12, False     ' family, size, bold, [italic]
'   PDF.PrintText "Dear John,", pdfLeft
'   PDF.PrintText ""                   ' blank line
'   PDF.PrintText "Please pay **$500.00** by Friday.", pdfJustify
'   PDF.NewPage                        ' force a page break
'   PDF.PrintText "Sincerely,", pdfRight
'   PDF.SavePDF
'
' Inline text styling (works in PrintText/PrintParagraph/PrintTabbedLine, and
' combines freely - e.g. "***bold italic***" or "**__bold underline__**"):
'   **bold**        - double asterisks toggle bold
'   *italic*        - single asterisk toggles italic
'   __underline__   - double underscores toggle underline (drawn as an actual
'                      line beneath the text, since PDF has no underline
'                      text attribute)
' Each toggles relative to the current SetFont Bold/Italic/Underline state.
' (A lone "*" or "_" occurring naturally in text - e.g. "3*4" - will be read
' as markup; use SetFont's Bold/Italic/Underline params instead if that's a
' concern for a given letter.)
' Alignment options: pdfLeft, pdfCenter, pdfRight, pdfJustify (the default).
' Justify stretches word spacing to reach the right margin; the final wrapped
' line of a paragraph is left-aligned, per normal typographic convention.
' Leading and multiple consecutive spaces are preserved as space-run tokens.
' With pdfJustify, a leading space-run indent stays fixed at the left and only
' the text after it is stretched to the right margin.
' vbNewLine / vbCrLf inside the string forces a line break; a blank line
' (vbNewLine & vbNewLine) inserts an empty row.
'
' Aliases (equivalent alternate spellings, for callers that prefer them):
'   SetMarginsPt Left, Top, Right, Bottom   - same as SetMargins but in points
'   SetFont "F1".."F12", Size               - direct PDF font-resource style
'                                              (F1=Helvetica ... F12=Courier-BoldOblique,
'                                              same mapping as GetBaseFontName below)
'   PrintParagraph Text, [Align]            - identical to PrintText
'
' Columns/tabs (for label/value rows or simple tables in the middle of a page):
'   PDF.SetTabStops 0, 2.5, 4.5             ' column start positions, inches from left margin
'   PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price", Array(pdfLeft, pdfRight, pdfRight)
'   PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**", Array(pdfLeft, pdfRight, pdfRight)
' Passing the same per-column Align array to both the header and data rows is
' what makes "Qty"/"Price" line up directly above the right-aligned numbers
' beneath them, while "Item" stays left-aligned.
' SetTabStopsPt is the points-based equivalent of SetTabStops.
'
' One-off horizontal placement (e.g. a signature line mid-page, without setting
' up tab stops): SetStartX applies only to the very next PrintText/PrintParagraph
' call's first line, then flow reverts to the normal left margin automatically.
'   PDF.SetStartX 3.5                       ' inches from the left margin
'   PDF.PrintText "X_______________"
' SetStartXPt is the points-based equivalent of SetStartX.
'
' Letterhead background (a pre-designed graphic - e.g. a firm's logo/heading -
' stamped behind the first page of every document, the way pre-printed letterhead
' paper works). The graphic itself comes from a one-time text extraction of a
' source PDF's drawing commands (see DOC\Extract_LH_Content.ps1), not the PDF
' file directly, since VB6 has no built-in Flate/zlib decompressor:
'   PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
'   PDF.SetLetterheadFooter "THE LAW FIRM OF ... IS A DEBT COLLECTOR ..."
'   PDF.PrintText "Dear John,"
'   ' ... rest of the letter ...
'   PDF.SavePDF
' Only the source PDF's page graphics are reproduced (no fonts/images), so this
' only works for letterhead artwork that is pure vector paths - exactly what
' DOC\LH.pdf (a "Print to PDF" vector letterhead) contains. ClearLetterhead turns
' it back off for any later documents created with the same clsPDF instance.
' NOTE: for a letterhead sidebar whose names/text need to change over time (e.g.
' an attorney/staff roster), prefer PrintTextAt/DrawLine below plus a real data
' source instead - see Letterhead_clsPDF.PrintFirmSidebar for a worked example
' that pulls its roster from the employee table.

' Absolute-position primitives (for content that runs independently of the
' normal top-down PrintText flow - e.g. a sidebar column down the side of a
' page while the main letter body flows down the left side). Unlike
' PrintText/SetStartX, these do NOT move the normal text cursor and never
' trigger a page break - they simply draw at the exact coordinates given,
' measured in inches (or points, for the *Pt variants) from the TOP-LEFT
' corner of the page (not from the margins):
'   PDF.DrawLine 6.15, 0.6, 6.15, 10.2              ' a vertical rule
'   PDF.PrintTextAt 6.35, 0.6, "Some Law Firm", pdfCenter, 1.65   ' centered in a 1.65in-wide box
'   PDF.PrintTextAt 1, 1, "Just a label"             ' WidthIn omitted -> plain left-aligned
' Align other than pdfLeft requires WidthIn/WidthPt (the box to align within);
' it's ignored (treated as pdfLeft) if omitted/zero. Both draw onto whichever
' page is "current" at the time of the call - call them before any NewPage if
' they're meant to land on page 1 (the usual case for a letterhead sidebar).

Public Enum PdfAlign
    pdfLeft = 0
    pdfCenter = 1
    pdfRight = 2
    pdfJustify = 3
End Enum

Private Type PdfToken
    sWord As String
    bBold As Boolean
    bItalic As Boolean
    bUnderline As Boolean
End Type

Private m_Started As Boolean
Private m_FilePath As String
Private m_PageWidth As Double
Private m_PageHeight As Double
Private m_MarginLeft As Double
Private m_MarginTop As Double
Private m_MarginRight As Double
Private m_MarginBottom As Double
Private m_CurX As Double
Private m_CurY As Double
Private m_FontFamily As String
Private m_FontSize As Double
Private m_FontBold As Boolean
Private m_FontItalic As Boolean
Private m_FontUnderline As Boolean
Private m_CurFontKey As String
Private m_CurContent As String
Private m_Pages() As String
Private m_PageCount As Integer

Private m_TabStops() As Double
Private m_TabStopCount As Integer

Private m_PendingStartX As Double
Private m_HasPendingStartX As Boolean

Private m_LetterheadContent As String
Private m_HasLetterhead As Boolean

Private m_LetterheadFooterText As String
Private m_LetterheadFooterY As Double   ' inches from top of page
Private m_HasLetterheadFooter As Boolean

Private m_WHelv(32 To 126) As Integer
Private m_WHelvB(32 To 126) As Integer
Private m_WTimes(32 To 126) As Integer
Private m_WTimesB(32 To 126) As Integer

Private Sub Class_Initialize()
    m_Started = False
    m_PageWidth = 612
    m_PageHeight = 792
    m_FontFamily = "Helvetica"
    m_FontSize = 12
    m_FontBold = False
    m_FontItalic = False
    m_FontUnderline = False
    m_TabStopCount = 0
    m_HasPendingStartX = False
    m_HasLetterhead = False
    m_HasLetterheadFooter = False
    LoadWidthTables
End Sub

'--------------------------------------------------------------------------
' Public API
'--------------------------------------------------------------------------

Public Sub CreatePDF(ByVal sFilePath As String)
    m_FilePath = sFilePath
    m_PageWidth = 612      ' US Letter, points (8.5in x 11in)
    m_PageHeight = 792
    m_PageCount = 0
    ReDim m_Pages(0 To 0)
    m_CurContent = ""
    m_CurFontKey = ""

    SetMargins 1, 1, 1, 1
    SetFont "Helvetica", 12, False, False

    m_Started = True
End Sub

' Returns the path last written by SavePDF (may differ from CreatePDF if the
' target file was locked open in a viewer and the PDF was saved to a .tmp file).
Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

Public Sub SetMargins(ByVal LeftIn As Double, ByVal TopIn As Double, ByVal RightIn As Double, ByVal BottomIn As Double)
    m_MarginLeft = LeftIn * 72
    m_MarginTop = TopIn * 72
    m_MarginRight = RightIn * 72
    m_MarginBottom = BottomIn * 72

    m_CurX = m_MarginLeft
    m_CurY = m_PageHeight - m_MarginTop
End Sub

' Same as SetMargins, but values are in points (1/72 inch) rather than inches.
Public Sub SetMarginsPt(ByVal LeftPt As Double, ByVal TopPt As Double, ByVal RightPt As Double, ByVal BottomPt As Double)
    m_MarginLeft = LeftPt
    m_MarginTop = TopPt
    m_MarginRight = RightPt
    m_MarginBottom = BottomPt

    m_CurX = m_MarginLeft
    m_CurY = m_PageHeight - m_MarginTop
End Sub

' Defines column positions (tab stops), measured in inches from the left
' margin, for use by PrintTabbedLine. Example: PDF.SetTabStops 0, 2.5, 4.5
' creates 3 columns starting at the margin, 2.5in in, and 4.5in in.
Public Sub SetTabStops(ParamArray Positions() As Variant)
    Dim vArr As Variant
    vArr = Positions
    ApplyTabStops vArr, 72
End Sub

' Same as SetTabStops, but values are in points (1/72 inch) rather than inches.
Public Sub SetTabStopsPt(ParamArray Positions() As Variant)
    Dim vArr As Variant
    vArr = Positions
    ApplyTabStops vArr, 1
End Sub

' One-time horizontal start position for the very next PrintText/PrintParagraph
' call only (applies to that call's first line, in case it word-wraps to more
' than one line - the wrapped lines return to the normal left margin, and every
' call after that reverts to normal flow automatically). Measured in inches
' from the left margin. Handy for a short label or signature line placed mid-
' page without setting up tab stops. Example:
'   PDF.SetStartX 3.5
'   PDF.PrintText "X_______________"
Public Sub SetStartX(ByVal XIn As Double)
    m_PendingStartX = XIn * 72
    m_HasPendingStartX = True
End Sub

' Same as SetStartX, but the value is in points (1/72 inch) rather than inches.
Public Sub SetStartXPt(ByVal XPt As Double)
    m_PendingStartX = XPt
    m_HasPendingStartX = True
End Sub

' Draws one line of text at an exact page position, independent of the normal
' top-down PrintText cursor (see the "Absolute-position primitives" note near
' the top of this file). XIn/YIn are inches from the page's top-left corner.
' If WidthIn is given (>0), Align (pdfLeft/pdfCenter/pdfRight/pdfJustify) positions
' the text within that box width, starting at XIn; otherwise the text is always
' left-aligned starting at XIn regardless of Align. Supports the same inline
' **bold**/*italic*/__underline__ markup as PrintText, using the current SetFont.
Public Sub PrintTextAt(ByVal XIn As Double, ByVal YIn As Double, ByVal Text As String, Optional ByVal Align As PdfAlign = pdfLeft, Optional ByVal WidthIn As Double = 0)
    PrintTextAtPt XIn * 72, YIn * 72, Text, Align, WidthIn * 72
End Sub

' Same as PrintTextAt, but values are in points (1/72 inch) rather than inches.
Public Sub PrintTextAtPt(ByVal XPt As Double, ByVal YPt As Double, ByVal Text As String, Optional ByVal Align As PdfAlign = pdfLeft, Optional ByVal WidthPt As Double = 0)
    Dim lineArr() As String
    Dim segIdx As Integer
    Dim lineNum As Integer
    Dim lineH As Double

    If Not m_Started Then
        Err.Raise 5, "clsPDF.PrintTextAt", "CreatePDF must be called before PrintTextAt."
    End If
    If Len(Text) = 0 Then Exit Sub

    lineArr = Split(NormalizeNewlines(Text), Chr(10))
    lineH = LineHeight()
    lineNum = 0

    For segIdx = 0 To UBound(lineArr)
        If Len(lineArr(segIdx)) = 0 Then
            lineNum = lineNum + 1
        Else
            lineNum = lineNum + PrintTextAtSegment(lineArr(segIdx), XPt, YPt, Align, WidthPt, lineNum, lineH)
        End If
    Next segIdx
End Sub

' Centers (or aligns) text within the full physical page width at the current
' vertical cursor — used for the FDCPA footer under the letterhead sidebar.
Public Sub PrintTextCenteredOnFullPage(ByVal Text As String, Optional ByVal Align As PdfAlign = pdfCenter)
    Dim yIn As Double
    Dim lineArr() As String
    Dim segIdx As Integer
    Dim lineNum As Integer
    Dim lineH As Double
    Dim rowsUsed As Integer
    Dim nRows As Integer

    If Not m_Started Then
        Err.Raise 5, "clsPDF.PrintTextCenteredOnFullPage", "CreatePDF must be called first."
    End If
    If Len(Text) = 0 Then Exit Sub

    If (m_CurY - LineHeight()) < m_MarginBottom Then NewPage

    yIn = (m_PageHeight - m_CurY) / 72
    lineH = LineHeight()
    lineArr = Split(NormalizeNewlines(Text), Chr(10))
    lineNum = 0
    rowsUsed = 0

    For segIdx = 0 To UBound(lineArr)
        If Len(lineArr(segIdx)) = 0 Then
            lineNum = lineNum + 1
            rowsUsed = rowsUsed + 1
        Else
            nRows = PrintTextAtSegment(lineArr(segIdx), 0, yIn * 72, Align, m_PageWidth, lineNum, lineH)
            lineNum = lineNum + nRows
            rowsUsed = rowsUsed + nRows
        End If
    Next segIdx

    m_CurY = m_CurY - rowsUsed * lineH
End Sub

' Draws a straight line at an exact page position, independent of the normal
' PrintText cursor (see the "Absolute-position primitives" note near the top
' of this file). Coordinates are inches from the page's top-left corner.
Public Sub DrawLine(ByVal X1In As Double, ByVal Y1In As Double, ByVal X2In As Double, ByVal Y2In As Double, Optional ByVal WidthPt As Double = 1)
    DrawLinePt X1In * 72, Y1In * 72, X2In * 72, Y2In * 72, WidthPt
End Sub

' Same as DrawLine, but coordinates are in points (1/72 inch) rather than inches.
Public Sub DrawLinePt(ByVal X1Pt As Double, ByVal Y1Pt As Double, ByVal X2Pt As Double, ByVal Y2Pt As Double, Optional ByVal WidthPt As Double = 1)
    If Not m_Started Then
        Err.Raise 5, "clsPDF.DrawLine", "CreatePDF must be called before DrawLine."
    End If
    AppendRaw "q" & vbLf
    AppendRaw Num(WidthPt) & " w" & vbLf
    AppendRaw Num(X1Pt) & " " & Num(m_PageHeight - Y1Pt) & " m" & vbLf
    AppendRaw Num(X2Pt) & " " & Num(m_PageHeight - Y2Pt) & " l" & vbLf
    AppendRaw "S" & vbLf
    AppendRaw "Q" & vbLf
End Sub

' Loads a pre-extracted letterhead graphic (raw PDF drawing commands, produced
' once via DOC\Extract_LH_Content.ps1 from a source PDF such as DOC\LH.pdf) and
' stamps it onto the first page of the next document(s) saved by this clsPDF
' instance. See the "Letterhead background" note near the top of this file.
Public Sub SetLetterhead(ByVal ContentFilePath As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open ContentFilePath For Input As #fnum
    m_LetterheadContent = Input$(LOF(fnum), fnum)
    Close #fnum
    m_HasLetterhead = True
End Sub

' Turns off the letterhead set by SetLetterhead for any documents created with
' this clsPDF instance from this point on.
Public Sub ClearLetterhead()
    m_HasLetterhead = False
End Sub

' Sets the bottom-of-page letterhead disclosure/footer text that is drawn as
' real text (not vector artwork) on page 1 when SavePDF runs. YFromTopIn is
' the vertical position of the first footer line, measured in inches from the
' top of the page (default 10.35" places a two-line footer near the bottom).
' Pass an empty string to suppress the footer without clearing the sidebar
' graphic loaded by SetLetterhead.
Public Sub SetLetterheadFooter(ByVal FooterText As String, Optional ByVal YFromTopIn As Double = 10.35)
    m_LetterheadFooterText = FooterText
    m_LetterheadFooterY = YFromTopIn
    m_HasLetterheadFooter = (Len(FooterText) > 0)
End Sub

Public Sub ClearLetterheadFooter()
    m_HasLetterheadFooter = False
    m_LetterheadFooterText = ""
End Sub

Public Sub NewPage()
    FlushCurrentPage
    m_CurContent = ""
    m_CurX = m_MarginLeft
    m_CurY = m_PageHeight - m_MarginTop
    m_CurFontKey = ""
End Sub

Public Sub SetFont(ByVal FontName As String, ByVal Size As Double, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal underline As Boolean = False)
    Select Case UCase$(Trim$(FontName))
        ' Direct PDF font-resource names (same mapping as GetBaseFontName)
        Case "F1": m_FontFamily = "Helvetica": bold = False: italic = False
        Case "F2": m_FontFamily = "Helvetica": bold = True: italic = False
        Case "F3": m_FontFamily = "Helvetica": bold = False: italic = True
        Case "F4": m_FontFamily = "Helvetica": bold = True: italic = True
        Case "F5": m_FontFamily = "Times": bold = False: italic = False
        Case "F6": m_FontFamily = "Times": bold = True: italic = False
        Case "F7": m_FontFamily = "Times": bold = False: italic = True
        Case "F8": m_FontFamily = "Times": bold = True: italic = True
        Case "F9": m_FontFamily = "Courier": bold = False: italic = False
        Case "F10": m_FontFamily = "Courier": bold = True: italic = False
        Case "F11": m_FontFamily = "Courier": bold = False: italic = True
        Case "F12": m_FontFamily = "Courier": bold = True: italic = True
        Case Else
            Select Case LCase$(FontName)
                Case "helvetica", "arial", "sans", "sans-serif"
                    m_FontFamily = "Helvetica"
                Case "times", "times new roman", "times-roman", "serif"
                    m_FontFamily = "Times"
                Case "courier", "courier new", "mono", "monospace"
                    m_FontFamily = "Courier"
                Case Else
                    m_FontFamily = "Helvetica"
            End Select
    End Select

    m_FontSize = Size
    m_FontBold = bold
    m_FontItalic = italic
    m_FontUnderline = underline
End Sub

Public Sub PrintText(ByVal Text As String, Optional ByVal Align As PdfAlign = pdfJustify)
    Dim lineArr() As String
    Dim segIdx As Integer
    Dim normalXStart As Double, normalAvailWidth As Double
    Dim segXStart As Double, segAvailWidth As Double

    If Not m_Started Then
        Err.Raise 5, "clsPDF.PrintText", "CreatePDF must be called before PrintText."
    End If

    normalXStart = m_MarginLeft
    normalAvailWidth = m_PageWidth - m_MarginLeft - m_MarginRight

    ' A pending SetStartX/SetStartXPt call only ever affects the first line of
    ' the very next PrintText/PrintParagraph call - consume it now so it can't
    ' leak into any call after this one.
    If m_HasPendingStartX Then
        segXStart = m_MarginLeft + m_PendingStartX
        m_HasPendingStartX = False
    Else
        segXStart = normalXStart
    End If
    segAvailWidth = m_PageWidth - segXStart - m_MarginRight

    If Len(Text) = 0 Then
        PrintTextBlankLine
        Exit Sub
    End If

    lineArr = Split(NormalizeNewlines(Text), Chr(10))
    For segIdx = 0 To UBound(lineArr)
        If segIdx > 0 Then
            segXStart = normalXStart
            segAvailWidth = normalAvailWidth
        End If

        If Len(lineArr(segIdx)) = 0 Then
            PrintTextBlankLine
        Else
            PrintTextSegment lineArr(segIdx), Align, segXStart, segAvailWidth, normalXStart, normalAvailWidth
        End If
    Next segIdx
End Sub

' Alias for PrintText.
Public Sub PrintParagraph(ByVal Text As String, Optional ByVal Align As PdfAlign = pdfJustify)
    PrintText Text, Align
End Sub

' Prints one row of columns. Text is split on Chr(9) (vbTab); each segment is
' placed at the corresponding column defined by SetTabStops/SetTabStopsPt (the
' 1st segment at the 1st tab stop, the 2nd at the 2nd, etc). Segments are not
' word-wrapped (columns are meant for short label/value or tabular data) but
' do support inline **bold** markup like PrintText. Align controls how each
' segment sits within its column, and can be either:
'   - a single PdfAlign, applied to every column (pdfLeft is the default), or
'   - Array(align1, align2, ...) to set each column's alignment individually
'     (handy for a text column that stays left-aligned alongside numeric
'     columns that are right-aligned - see the header/data rows below, where
'     "Qty"/"Price" line up directly over the right-aligned numbers beneath
'     them). If fewer aligns than columns are given, the last one repeats for
'     any remaining columns.
' If more segments are supplied than tab stops are defined, extra segments
' simply continue in natural flow.
' Example:
'   PDF.SetTabStops 0, 2.5, 4.5
'   PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price", Array(pdfLeft, pdfRight, pdfRight)
'   PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**", Array(pdfLeft, pdfRight, pdfRight)
Public Sub PrintTabbedLine(ByVal Text As String, Optional ByVal Align As Variant = pdfLeft)
    Dim segments() As String
    Dim nSeg As Integer
    Dim I As Integer
    Dim colX As Double, colWidth As Double, rightEdge As Double
    Dim lastX As Double
    Dim segAlign As PdfAlign

    If Not m_Started Then
        Err.Raise 5, "clsPDF.PrintTabbedLine", "CreatePDF must be called before PrintTabbedLine."
    End If

    If (m_CurY - LineHeight()) < m_MarginBottom Then NewPage

    segments = Split(Text, vbTab)
    nSeg = UBound(segments) + 1
    lastX = m_MarginLeft

    For I = 0 To nSeg - 1
        If I < m_TabStopCount Then
            colX = m_MarginLeft + m_TabStops(I)
        Else
            colX = lastX + 4 ' no more tab stops defined - keep flowing with a small gap
        End If

        If I + 1 < m_TabStopCount Then
            rightEdge = m_MarginLeft + m_TabStops(I + 1)
        Else
            rightEdge = m_PageWidth - m_MarginRight
        End If

        colWidth = rightEdge - colX
        If colWidth < 0 Then colWidth = 0

        segAlign = ColumnAlign(Align, I)
        lastX = EmitColumnSegment(segments(I), colX, colWidth, segAlign)
    Next I

    m_CurY = m_CurY - LineHeight()
End Sub

' Resolves the PdfAlign to use for column index ColIdx (0-based) from the
' Align argument passed to PrintTabbedLine, which may be a single PdfAlign
' (used for every column) or an array of PdfAlign values (one per column;
' the last element repeats for any columns beyond the array's length).
Private Function ColumnAlign(ByVal Align As Variant, ByVal ColIdx As Integer) As PdfAlign
    If IsArray(Align) Then
        If ColIdx > UBound(Align) Then
            ColumnAlign = CLng(Align(UBound(Align)))
        Else
            ColumnAlign = CLng(Align(ColIdx))
        End If
    Else
        ColumnAlign = CLng(Align)
    End If
End Function

Public Sub SavePDF()
    Dim nPages As Integer
    Dim I As Integer, k As Integer
    Dim totalObjs As Integer
    Dim offsets() As Long
    Dim pageObjNum As Integer, contObjNum As Integer, fontObjNum As Integer
    Dim resName As String, baseName As String
    Dim resourcesDict As String
    Dim streamData As String
    Dim fnum As Integer
    Dim Pos As Long
    Dim xrefStart As Long
    Dim xrefStr As String
    Dim kids As String
    Dim letterheadObjNum As Integer
    Dim lhStream As String
    Dim contentsEntry As String
    Dim tempPath As String
    Dim targetPath As String

    If Not m_Started Then
        Err.Raise 5, "clsPDF.SavePDF", "CreatePDF must be called before SavePDF."
    End If

    targetPath = m_FilePath
    tempPath = BuildTempPdfPath(targetPath)

    FlushCurrentPage
    nPages = m_PageCount
    If nPages = 0 Then
        NewPage
        FlushCurrentPage
        nPages = m_PageCount
    End If

    If m_HasLetterheadFooter And nPages >= 1 Then
        m_Pages(0) = BuildLetterheadFooterContent() & m_Pages(0)
    End If

    totalObjs = 14 + nPages * 2
    If m_HasLetterhead Then
        letterheadObjNum = totalObjs + 1
        totalObjs = totalObjs + 1
    End If
    ReDim offsets(1 To totalObjs)

    On Error Resume Next
    If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    On Error GoTo 0

    fnum = FreeFile
    On Error GoTo SavePdfWriteFailed
    Open tempPath For Binary Access Write As #fnum
    On Error GoTo 0
    Pos = 1

    WriteBytes fnum, Pos, "%PDF-1.4" & vbLf & "%" & Chr$(226) & Chr$(227) & Chr$(207) & Chr$(211) & vbLf

    ' Object 1: Catalog
    offsets(1) = Pos
    WriteBytes fnum, Pos, "1 0 obj" & vbLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbLf & "endobj" & vbLf

    ' Object 2: Pages
    kids = ""
    For I = 1 To nPages
        kids = kids & (15 + (I - 1) * 2) & " 0 R "
    Next I
    offsets(2) = Pos
    WriteBytes fnum, Pos, "2 0 obj" & vbLf & "<< /Type /Pages /Kids [" & Trim$(kids) & "] /Count " & nPages & " >>" & vbLf & "endobj" & vbLf

    ' Objects 3-14: the 12 standard base-14 fonts we support (no embedding needed)
    For fontObjNum = 3 To 14
        resName = "F" & (fontObjNum - 2)
        baseName = GetBaseFontName(resName)
        offsets(fontObjNum) = Pos
        WriteBytes fnum, Pos, fontObjNum & " 0 obj" & vbLf & "<< /Type /Font /Subtype /Type1 /BaseFont /" & baseName & " /Encoding /WinAnsiEncoding >>" & vbLf & "endobj" & vbLf
    Next fontObjNum

    resourcesDict = "<< /Font << "
    For k = 1 To 12
        resourcesDict = resourcesDict & "/F" & k & " " & (k + 2) & " 0 R "
    Next k
    resourcesDict = resourcesDict & ">> >>"

    ' Letterhead object (if set): one shared stream, referenced from page 1's
    ' /Contents only, wrapped in its own q/Q so it can't leak graphics state
    ' (CTM, color, line width, ...) into that page's own content afterward.
    If m_HasLetterhead Then
        lhStream = "q" & vbLf & m_LetterheadContent & vbLf & "Q" & vbLf
        offsets(letterheadObjNum) = Pos
        WriteBytes fnum, Pos, letterheadObjNum & " 0 obj" & vbLf & "<< /Length " & Len(lhStream) & " >>" & vbLf & "stream" & vbLf & lhStream & "endstream" & vbLf & "endobj" & vbLf
    End If

    ' Page + content-stream object pairs
    For I = 1 To nPages
        pageObjNum = 15 + (I - 1) * 2
        contObjNum = pageObjNum + 1

        If m_HasLetterhead And I = 1 Then
            contentsEntry = "[" & letterheadObjNum & " 0 R " & contObjNum & " 0 R]"
        Else
            contentsEntry = contObjNum & " 0 R"
        End If

        offsets(pageObjNum) = Pos
        WriteBytes fnum, Pos, pageObjNum & " 0 obj" & vbLf & _
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & Num(m_PageWidth) & " " & Num(m_PageHeight) & "] /Resources " & resourcesDict & " /Contents " & contentsEntry & " >>" & vbLf & "endobj" & vbLf

        streamData = m_Pages(I - 1)
        offsets(contObjNum) = Pos
        WriteBytes fnum, Pos, contObjNum & " 0 obj" & vbLf & "<< /Length " & Len(streamData) & " >>" & vbLf & "stream" & vbLf & streamData & "endstream" & vbLf & "endobj" & vbLf
    Next I

    ' xref table - each entry must be exactly 20 bytes (10-digit offset, 5-digit gen, type, CRLF)
    xrefStart = Pos
    xrefStr = "xref" & vbLf & "0 " & (totalObjs + 1) & vbLf & "0000000000 65535 f" & vbCrLf
    For I = 1 To totalObjs
        xrefStr = xrefStr & Right$("0000000000" & CStr(offsets(I)), 10) & " 00000 n" & vbCrLf
    Next I
    WriteBytes fnum, Pos, xrefStr

    WriteBytes fnum, Pos, "trailer" & vbLf & "<< /Size " & (totalObjs + 1) & " /Root 1 0 R >>" & vbLf & "startxref" & vbLf & xrefStart & vbLf & "%%EOF"

    Close #fnum

    If CommitPdfFile(tempPath, targetPath) Then
        m_FilePath = targetPath
    Else
        m_FilePath = tempPath
    End If
    m_Started = False
    Exit Sub

SavePdfWriteFailed:
    Dim writeErr As Long
    Dim writeMsg As String
    writeErr = Err.Number
    writeMsg = Err.Description
    On Error Resume Next
    Close #fnum
    On Error GoTo 0
    Err.Raise 75, "clsPDF.SavePDF", "Cannot write PDF file (" & tempPath & "). " & writeMsg & " (" & writeErr & ")"
End Sub

Private Function BuildTempPdfPath(ByVal targetPath As String) As String
    Dim dotPos As Long

    dotPos = InStrRev(targetPath, ".")
    If dotPos > InStrRev(targetPath, "\") Then
        BuildTempPdfPath = Left$(targetPath, dotPos - 1) & ".tmp"
    Else
        BuildTempPdfPath = targetPath & ".tmp"
    End If
End Function

' Write to a .tmp file first, then replace the target. Returns False when the
' target is locked (usually because a PDF viewer still has it open).
Private Function CommitPdfFile(ByVal tempPath As String, ByVal targetPath As String) As Boolean
    On Error GoTo CommitFailed

    If StrComp(tempPath, targetPath, vbTextCompare) = 0 Then
        CommitPdfFile = True
        Exit Function
    End If

    If Len(Dir$(targetPath)) > 0 Then
        On Error Resume Next
        SetAttr targetPath, vbNormal
        On Error GoTo CommitFailed
        Kill targetPath
    End If
    Name tempPath As targetPath
    CommitPdfFile = True
    Exit Function

CommitFailed:
    CommitPdfFile = False
End Function

'--------------------------------------------------------------------------
' Internal: line layout (word-wrap, alignment, justification)
'--------------------------------------------------------------------------

Private Function LeadingSpacePrefixCount(ByRef tokens() As PdfToken, ByVal nCount As Integer) As Integer
    Dim I As Integer

    LeadingSpacePrefixCount = 0
    For I = 0 To nCount - 1
        If IsSpaceRun(tokens(I).sWord) Then
            LeadingSpacePrefixCount = LeadingSpacePrefixCount + 1
        Else
            Exit For
        End If
    Next I
End Function

Private Function MeasureTokenSliceWidth(ByRef tokens() As PdfToken, ByVal startIdx As Integer, ByVal endIdx As Integer, ByRef tokW() As Double) As Double
    Dim I As Integer
    Dim spaceW As Double
    Dim totalW As Double

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    totalW = 0
    For I = startIdx To endIdx
        totalW = totalW + tokW(I)
        If I < endIdx Then
            If NeedsInterWordSpace(tokens, I) Then totalW = totalW + spaceW
        End If
    Next I
    MeasureTokenSliceWidth = totalW
End Function

Private Sub EmitLine(ByRef tokens() As PdfToken, ByVal nCount As Integer, ByVal xStart0 As Double, ByVal availWidth As Double, ByVal Align As PdfAlign, ByVal isLast As Boolean)
    Dim naturalWidth As Double
    Dim spaceW As Double
    Dim I As Integer
    Dim xStart As Double
    Dim tw As Double
    Dim tokW() As Double
    Dim tokX() As Double
    Dim pen As Double
    Dim nGaps As Integer
    Dim nLead As Integer
    Dim indentW As Double
    Dim bodyStart As Integer
    Dim bodyAvail As Double
    Dim bodyNatural As Double
    Dim justifyBody As Boolean

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    naturalWidth = MeasureTokenRowWidth(tokens, nCount, tokW)

    nLead = LeadingSpacePrefixCount(tokens, nCount)
    indentW = 0
    If nLead > 0 Then indentW = MeasureTokenSliceWidth(tokens, 0, nLead - 1, tokW)
    bodyStart = nLead

    tw = 0
    Select Case Align
        Case pdfCenter
            xStart = xStart0 + (availWidth - naturalWidth) / 2
        Case pdfRight
            xStart = xStart0 + (availWidth - naturalWidth)
        Case pdfJustify
            xStart = xStart0
            If nCount - bodyStart > 0 Then
                nGaps = CountWordToWordGapsSlice(tokens, bodyStart, nCount - 1)
                bodyNatural = MeasureTokenSliceWidth(tokens, bodyStart, nCount - 1, tokW)
                bodyAvail = availWidth - indentW
                justifyBody = (Not isLast)
                If justifyBody And nGaps > 0 Then tw = (bodyAvail - bodyNatural) / nGaps
            End If
        Case Else ' pdfLeft
            xStart = xStart0
    End Select

    If (m_CurY - LineHeight()) < m_MarginBottom Then NewPage

    ReDim tokX(0 To nCount - 1)
    If Align = pdfJustify And nLead > 0 Then
        pen = xStart0
        For I = 0 To nLead - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, pdfLeft, 0, spaceW
        Next I
        pen = xStart0 + indentW
        For I = bodyStart To nCount - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, pdfJustify, tw, spaceW
        Next I
    Else
        pen = xStart
        For I = 0 To nCount - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, Align, tw, spaceW
        Next I
    End If

    EmitTokensAtRow tokens, nCount, m_CurY, tokW, tokX

    EmitUnderlines tokens, tokX, tokW, nCount

    m_CurY = m_CurY - LineHeight()
End Sub

' Renders tokens at pre-computed X positions using absolute Tm for each token.
' Avoids PDF Tw, which would stretch every space character (including leading
' indents and intentional multi-space runs) during justified output.
Private Sub EmitTokensAtRow(ByRef tokens() As PdfToken, ByVal nCount As Integer, ByVal yDev As Double, ByRef tokW() As Double, ByRef tokX() As Double)
    Dim I As Integer

    AppendRaw "BT" & vbLf
    AppendRaw "0 Tw" & vbLf
    For I = 0 To nCount - 1
        SetFontIfNeeded tokens(I).bBold, tokens(I).bItalic
        AppendRaw "1 0 0 1 " & Num(tokX(I)) & " " & Num(yDev) & " Tm" & vbLf
        AppendRaw "(" & EscapePdfString(tokens(I).sWord) & ") Tj" & vbLf
    Next I
    AppendRaw "ET" & vbLf
End Sub

' Like EmitLine, but draws at an explicit device-space Y (yDev, already flipped
' to PDF's bottom-up coordinate system) instead of the running m_CurY cursor,
' and never advances that cursor or triggers a page break - used by
' PrintTextAt for content that runs independently of the normal text flow
' (e.g. a letterhead sidebar). If availWidth is 0, Align is always treated as
' pdfLeft (there's no box to align within).
Private Function NeedsInterWordSpaceTokens(ByRef leftTok As PdfToken, ByRef rightTok As PdfToken) As Boolean
    NeedsInterWordSpaceTokens = Not (IsSpaceRun(leftTok.sWord) Or IsSpaceRun(rightTok.sWord))
End Function

Private Sub EmitLineAt(ByRef tokens() As PdfToken, ByVal nCount As Integer, ByVal xStart0 As Double, ByVal yDev As Double, ByVal Align As PdfAlign, ByVal availWidth As Double)
    Dim naturalWidth As Double
    Dim spaceW As Double
    Dim I As Integer
    Dim xStart As Double
    Dim tw As Double
    Dim tokW() As Double
    Dim tokX() As Double
    Dim pen As Double
    Dim effAlign As PdfAlign
    Dim nGaps As Integer
    Dim nLead As Integer
    Dim indentW As Double
    Dim bodyStart As Integer
    Dim bodyAvail As Double
    Dim bodyNatural As Double

    effAlign = Align
    If availWidth <= 0 Then effAlign = pdfLeft

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    naturalWidth = MeasureTokenRowWidth(tokens, nCount, tokW)

    nLead = LeadingSpacePrefixCount(tokens, nCount)
    indentW = 0
    If nLead > 0 Then indentW = MeasureTokenSliceWidth(tokens, 0, nLead - 1, tokW)
    bodyStart = nLead

    tw = 0
    Select Case effAlign
        Case pdfCenter
            xStart = xStart0 + (availWidth - naturalWidth) / 2
        Case pdfRight
            xStart = xStart0 + (availWidth - naturalWidth)
        Case pdfJustify
            xStart = xStart0
            If nCount - bodyStart > 0 Then
                nGaps = CountWordToWordGapsSlice(tokens, bodyStart, nCount - 1)
                bodyNatural = MeasureTokenSliceWidth(tokens, bodyStart, nCount - 1, tokW)
                bodyAvail = availWidth - indentW
                If nGaps > 0 Then tw = (bodyAvail - bodyNatural) / nGaps
            End If
        Case Else ' pdfLeft
            xStart = xStart0
    End Select

    ReDim tokX(0 To nCount - 1)
    If effAlign = pdfJustify And nLead > 0 Then
        pen = xStart0
        For I = 0 To nLead - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, pdfLeft, 0, spaceW
        Next I
        pen = xStart0 + indentW
        For I = bodyStart To nCount - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, pdfJustify, tw, spaceW
        Next I
    Else
        pen = xStart
        For I = 0 To nCount - 1
            tokX(I) = pen
            pen = pen + tokW(I)
            AdvancePenAfterToken pen, tokens, I, nCount, effAlign, tw, spaceW
        Next I
    End If

    EmitTokensAtRow tokens, nCount, yDev, tokW, tokX

    EmitUnderlinesAt tokens, tokX, tokW, nCount, yDev
End Sub

' Draws a solid bar beneath each run of consecutive underlined tokens (so
' "__two words__" gets one continuous underline, not two separate ones).
Private Sub EmitUnderlines(ByRef tokens() As PdfToken, ByRef tokX() As Double, ByRef tokW() As Double, ByVal nCount As Integer)
    EmitUnderlinesAt tokens, tokX, tokW, nCount, m_CurY
End Sub

Private Sub EmitUnderlinesAt(ByRef tokens() As PdfToken, ByRef tokX() As Double, ByRef tokW() As Double, ByVal nCount As Integer, ByVal yDev As Double)
    Dim I As Integer
    Dim runStart As Integer
    Dim uStart As Double, uEnd As Double
    Dim underlineY As Double, underlineH As Double

    underlineY = yDev - m_FontSize * 0.08
    underlineH = m_FontSize * 0.06
    If underlineH < 0.75 Then underlineH = 0.75

    runStart = -1
    For I = 0 To nCount - 1
        If tokens(I).bUnderline Then
            If runStart = -1 Then runStart = I
            If I = nCount - 1 Then
                uStart = tokX(runStart)
                uEnd = tokX(I) + tokW(I)
                AppendRaw Num(uStart) & " " & Num(underlineY) & " " & Num(uEnd - uStart) & " " & Num(underlineH) & " re f" & vbLf
                runStart = -1
            End If
        Else
            If runStart <> -1 Then
                uStart = tokX(runStart)
                uEnd = tokX(I - 1) + tokW(I - 1)
                AppendRaw Num(uStart) & " " & Num(underlineY) & " " & Num(uEnd - uStart) & " " & Num(underlineH) & " re f" & vbLf
                runStart = -1
            End If
        End If
    Next I
End Sub

' Renders one column segment at a fixed X position (used by PrintTabbedLine).
' Does not word-wrap; returns the X position where the segment's text ended
' (colX + its natural width), used to keep overflow segments flowing sanely.
Private Function EmitColumnSegment(ByVal segText As String, ByVal colX As Double, ByVal colWidth As Double, ByVal Align As PdfAlign) As Double
    Dim tokens() As PdfToken
    Dim nTok As Integer
    Dim naturalWidth As Double
    Dim spaceW As Double
    Dim I As Integer
    Dim xStart As Double
    Dim tokW() As Double
    Dim tokX() As Double
    Dim pen As Double

    EmitColumnSegment = colX

    TokenizeLine segText, tokens, nTok
    If nTok = 0 Then Exit Function

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    naturalWidth = MeasureTokenRowWidth(tokens, nTok, tokW)

    Select Case Align
        Case pdfCenter
            xStart = colX + (colWidth - naturalWidth) / 2
        Case pdfRight
            xStart = colX + (colWidth - naturalWidth)
        Case Else ' pdfLeft / pdfJustify (justify isn't meaningful for a single column cell)
            xStart = colX
    End Select

    ReDim tokX(0 To nTok - 1)
    pen = xStart
    For I = 0 To nTok - 1
        tokX(I) = pen
        pen = pen + tokW(I)
        If I < nTok - 1 Then
            If NeedsInterWordSpace(tokens, I) Then pen = pen + spaceW
        End If
    Next I

    EmitTokensAtRow tokens, nTok, m_CurY, tokW, tokX

    EmitUnderlines tokens, tokX, tokW, nTok

    EmitColumnSegment = xStart + naturalWidth
End Function

Private Function NormalizeNewlines(ByVal sText As String) As String
    Dim r As String

    r = Replace$(sText, vbCrLf, Chr(10))
    r = Replace$(r, vbCr, Chr(10))
    r = Replace$(r, vbLf, Chr(10))
    NormalizeNewlines = r
End Function

Private Sub PrintTextBlankLine()
    If (m_CurY - LineHeight()) < m_MarginBottom Then NewPage
    m_CurY = m_CurY - LineHeight()
End Sub

' Word-wraps and emits one segment (no embedded newlines) in the normal flow.
Private Sub PrintTextSegment(ByVal Text As String, ByVal Align As PdfAlign, ByVal segXStart As Double, ByVal segAvailWidth As Double, ByVal normalXStart As Double, ByVal normalAvailWidth As Double)
    Dim tokens() As PdfToken
    Dim nTok As Integer
    Dim lineTok() As PdfToken
    Dim lineCount As Integer
    Dim I As Integer
    Dim wordW As Double, spaceW As Double, curWidth As Double, testWidth As Double
    Dim curXStart As Double, curAvailWidth As Double
    Dim paraXStart As Double, paraAvailWidth As Double

    paraXStart = segXStart
    paraAvailWidth = segAvailWidth
    curXStart = segXStart
    curAvailWidth = segAvailWidth

    TokenizeLine Text, tokens, nTok
    If nTok = 0 Then Exit Sub

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000

    ReDim lineTok(0 To nTok - 1)
    lineCount = 0
    curWidth = 0

    I = 0
    Do While I <= nTok - 1
        If lineCount = 0 And I > 0 And IsSpaceRun(tokens(I).sWord) Then
            I = I + 1
        Else
            wordW = MeasureWord(tokens(I).sWord, m_FontFamily, tokens(I).bBold) * m_FontSize / 1000

            If lineCount = 0 Then
                testWidth = wordW
            Else
                If NeedsInterWordSpaceTokens(lineTok(lineCount - 1), tokens(I)) Then
                    testWidth = curWidth + spaceW + wordW
                Else
                    testWidth = curWidth + wordW
                End If
            End If

            If testWidth > curAvailWidth And lineCount > 0 Then
                lineCount = TrimTrailingSpaceRuns(lineTok, lineCount)
                If lineCount > 0 Then
                    EmitLine lineTok, lineCount, curXStart, curAvailWidth, Align, False
                End If
                ' SetStartX applies to the first line only; wrapped lines use the
                ' normal left margin so the paragraph body spans the full width.
                curXStart = normalXStart
                curAvailWidth = normalAvailWidth
                lineCount = 0
                curWidth = 0
                If IsSpaceRun(tokens(I).sWord) Then
                    I = I + 1
                Else
                    lineTok(lineCount) = tokens(I)
                    lineCount = 1
                    curWidth = wordW
                    I = I + 1
                End If
            Else
                curWidth = testWidth
                lineTok(lineCount) = tokens(I)
                lineCount = lineCount + 1
                I = I + 1
            End If
        End If
    Loop

    lineCount = TrimTrailingSpaceRuns(lineTok, lineCount)
    If lineCount > 0 Then EmitLine lineTok, lineCount, curXStart, curAvailWidth, Align, True
End Sub

' Word-wraps and emits one segment at a fixed page position; returns the
' number of text rows used (including wrapped lines).
Private Function PrintTextAtSegment(ByVal Text As String, ByVal XPt As Double, ByVal YPt As Double, ByVal Align As PdfAlign, ByVal WidthPt As Double, ByVal lineNum As Integer, ByVal lineH As Double) As Integer
    Dim tokens() As PdfToken
    Dim nTok As Integer
    Dim lineTok() As PdfToken
    Dim lineCount As Integer
    Dim I As Integer
    Dim wordW As Double, spaceW As Double, curWidth As Double, testWidth As Double
    Dim segX As Double, segW As Double
    Dim rowsUsed As Integer
    Dim yDev As Double

    segX = XPt
    segW = WidthPt

    TokenizeLine Text, tokens, nTok
    If nTok = 0 Then
        PrintTextAtSegment = 0
        Exit Function
    End If

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    ReDim lineTok(0 To nTok - 1)
    lineCount = 0
    curWidth = 0
    rowsUsed = 0

    I = 0
    Do While I <= nTok - 1
        If lineCount = 0 And I > 0 And IsSpaceRun(tokens(I).sWord) Then
            I = I + 1
        Else
            wordW = MeasureWord(tokens(I).sWord, m_FontFamily, tokens(I).bBold) * m_FontSize / 1000

            If lineCount = 0 Then
                testWidth = wordW
            Else
                If NeedsInterWordSpaceTokens(lineTok(lineCount - 1), tokens(I)) Then
                    testWidth = curWidth + spaceW + wordW
                Else
                    testWidth = curWidth + wordW
                End If
            End If

            If segW > 0 Then
                If testWidth > segW And lineCount > 0 Then
                    lineCount = TrimTrailingSpaceRuns(lineTok, lineCount)
                    If lineCount > 0 Then
                        yDev = m_PageHeight - (YPt + (lineNum + rowsUsed) * lineH)
                        EmitLineAt lineTok, lineCount, segX, yDev, Align, segW
                        rowsUsed = rowsUsed + 1
                    End If
                    lineCount = 0
                    curWidth = 0
                    If IsSpaceRun(tokens(I).sWord) Then
                        I = I + 1
                    Else
                        lineTok(lineCount) = tokens(I)
                        lineCount = 1
                        curWidth = wordW
                        I = I + 1
                    End If
                Else
                    curWidth = testWidth
                    lineTok(lineCount) = tokens(I)
                    lineCount = lineCount + 1
                    I = I + 1
                End If
            Else
                curWidth = testWidth
                lineTok(lineCount) = tokens(I)
                lineCount = lineCount + 1
                I = I + 1
            End If
        End If
    Loop

    lineCount = TrimTrailingSpaceRuns(lineTok, lineCount)
    If lineCount > 0 Then
        yDev = m_PageHeight - (YPt + (lineNum + rowsUsed) * lineH)
        EmitLineAt lineTok, lineCount, segX, yDev, Align, segW
        rowsUsed = rowsUsed + 1
    End If

    PrintTextAtSegment = rowsUsed
End Function

Private Function TrimTrailingSpaceRuns(ByRef tokens() As PdfToken, ByVal nCount As Integer) As Integer
    Do While nCount > 0 And IsSpaceRun(tokens(nCount - 1).sWord)
        nCount = nCount - 1
    Loop
    TrimTrailingSpaceRuns = nCount
End Function

Private Function IsSpaceRun(ByVal sWord As String) As Boolean
    Dim I As Long

    If Len(sWord) = 0 Then
        IsSpaceRun = False
        Exit Function
    End If
    For I = 1 To Len(sWord)
        If mid$(sWord, I, 1) <> " " Then
            IsSpaceRun = False
            Exit Function
        End If
    Next I
    IsSpaceRun = True
End Function

' True when EmitLine should insert a normal single-word gap between two tokens.
Private Function NeedsInterWordSpace(ByRef tokens() As PdfToken, ByVal idx As Integer) As Boolean
    NeedsInterWordSpace = Not (IsSpaceRun(tokens(idx).sWord) Or IsSpaceRun(tokens(idx + 1).sWord))
End Function

' A space-run token that sits between two word tokens (the usual inter-word gap).
Private Function IsInterWordSpaceToken(ByRef tokens() As PdfToken, ByVal nCount As Integer, ByVal idx As Integer) As Boolean
    If Not IsSpaceRun(tokens(idx).sWord) Then
        IsInterWordSpaceToken = False
        Exit Function
    End If
    If idx <= 0 Or idx >= nCount - 1 Then
        IsInterWordSpaceToken = False
        Exit Function
    End If
    IsInterWordSpaceToken = (Not IsSpaceRun(tokens(idx - 1).sWord)) And (Not IsSpaceRun(tokens(idx + 1).sWord))
End Function

' Count word-to-word boundaries, skipping any space-run tokens between words.
Private Function CountWordToWordGapsSlice(ByRef tokens() As PdfToken, ByVal startIdx As Integer, ByVal endIdx As Integer) As Integer
    Dim I As Integer
    Dim J As Integer
    Dim nGaps As Integer

    nGaps = 0
    For I = startIdx To endIdx
        If Not IsSpaceRun(tokens(I).sWord) Then
            For J = I + 1 To endIdx
                If Not IsSpaceRun(tokens(J).sWord) Then
                    nGaps = nGaps + 1
                    Exit For
                End If
            Next J
        End If
    Next I
    CountWordToWordGapsSlice = nGaps
End Function

Private Sub AdvancePenAfterToken(ByRef pen As Double, ByRef tokens() As PdfToken, ByVal idx As Integer, ByVal nCount As Integer, ByVal Align As PdfAlign, ByVal tw As Double, ByVal spaceW As Double)
    If idx >= nCount - 1 Then Exit Sub

    If Align = pdfJustify And tw <> 0 Then
        If IsInterWordSpaceToken(tokens, nCount, idx) Then
            pen = pen + tw
        ElseIf NeedsInterWordSpace(tokens, idx) Then
            pen = pen + spaceW + tw
        End If
    ElseIf NeedsInterWordSpace(tokens, idx) Then
        pen = pen + spaceW
    End If
End Sub

Private Function MeasureTokenRowWidth(ByRef tokens() As PdfToken, ByVal nCount As Integer, ByRef tokW() As Double) As Double
    Dim I As Integer
    Dim spaceW As Double
    Dim totalW As Double

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    ReDim tokW(0 To nCount - 1)
    totalW = 0
    For I = 0 To nCount - 1
        tokW(I) = MeasureWord(tokens(I).sWord, m_FontFamily, tokens(I).bBold) * m_FontSize / 1000
        totalW = totalW + tokW(I)
        If I < nCount - 1 Then
            If NeedsInterWordSpace(tokens, I) Then totalW = totalW + spaceW
        End If
    Next I
    MeasureTokenRowWidth = totalW
End Function

Private Sub TokenizeLine(ByVal sText As String, ByRef tokens() As PdfToken, ByRef nCount As Integer)
    Dim I As Long, L As Long
    Dim bBold As Boolean, bItalic As Boolean, bUnderline As Boolean
    Dim bWordBold As Boolean, bWordItalic As Boolean, bWordUnderline As Boolean
    Dim sCurWord As String
    Dim sChar As String

    nCount = 0
    ReDim tokens(0 To 0)
    bBold = m_FontBold
    bItalic = m_FontItalic
    bUnderline = m_FontUnderline
    bWordBold = bBold
    bWordItalic = bItalic
    bWordUnderline = bUnderline
    sCurWord = ""
    L = Len(sText)
    I = 1
    Do While I <= L
        If mid$(sText, I, 2) = "**" Then
            bBold = Not bBold
            I = I + 2
        ElseIf mid$(sText, I, 2) = "__" Then
            bUnderline = Not bUnderline
            I = I + 2
        ElseIf mid$(sText, I, 1) = "*" Then
            bItalic = Not bItalic
            I = I + 1
        Else
            sChar = mid$(sText, I, 1)
            If sChar = " " Then
                If Len(sCurWord) > 0 Then
                    If IsSpaceRun(sCurWord) Then
                        sCurWord = sCurWord & " "
                    Else
                        AddToken tokens, nCount, sCurWord, bWordBold, bWordItalic, bWordUnderline
                        sCurWord = " "
                        bWordBold = bBold
                        bWordItalic = bItalic
                        bWordUnderline = bUnderline
                    End If
                Else
                    sCurWord = " "
                    bWordBold = bBold
                    bWordItalic = bItalic
                    bWordUnderline = bUnderline
                End If
            Else
                If Len(sCurWord) > 0 And IsSpaceRun(sCurWord) Then
                    AddToken tokens, nCount, sCurWord, bWordBold, bWordItalic, bWordUnderline
                    sCurWord = ""
                End If
                If Len(sCurWord) = 0 Then
                    bWordBold = bBold
                    bWordItalic = bItalic
                    bWordUnderline = bUnderline
                End If
                sCurWord = sCurWord & sChar
            End If
            I = I + 1
        End If
    Loop
    If Len(sCurWord) > 0 Then AddToken tokens, nCount, sCurWord, bWordBold, bWordItalic, bWordUnderline
End Sub

Private Sub AddToken(ByRef tokens() As PdfToken, ByRef nCount As Integer, ByVal sWord As String, ByVal bBold As Boolean, ByVal bItalic As Boolean, ByVal bUnderline As Boolean)
    ReDim Preserve tokens(0 To nCount)
    tokens(nCount).sWord = sWord
    tokens(nCount).bBold = bBold
    tokens(nCount).bItalic = bItalic
    tokens(nCount).bUnderline = bUnderline
    nCount = nCount + 1
End Sub

Private Function LineHeight() As Double
    LineHeight = m_FontSize * 1.2
End Function

'--------------------------------------------------------------------------
' Internal: font metrics (hardcoded AFM widths, 1/1000 em, for the base-14
' fonts that render without embedding). Oblique/Italic variants reuse their
' upright counterpart's widths as a close approximation.
'--------------------------------------------------------------------------

Private Function MeasureChar(ByVal ch As String, ByVal family As String, ByVal bold As Boolean) As Integer
    Dim code As Integer

    If family = "Courier" Then
        MeasureChar = 600
        Exit Function
    End If

    code = Asc(ch)
    If code < 32 Or code > 126 Then
        MeasureChar = 500
        Exit Function
    End If

    If family = "Times" Then
        If bold Then
            MeasureChar = m_WTimesB(code)
        Else
            MeasureChar = m_WTimes(code)
        End If
    Else ' Helvetica
        If bold Then
            MeasureChar = m_WHelvB(code)
        Else
            MeasureChar = m_WHelv(code)
        End If
    End If
End Function

Private Function MeasureWord(ByVal word As String, ByVal family As String, ByVal bold As Boolean) As Double
    Dim I As Long, total As Long
    total = 0
    For I = 1 To Len(word)
        total = total + MeasureChar(mid$(word, I, 1), family, bold)
    Next I
    MeasureWord = total
End Function

Private Sub LoadWidthTables()
    Dim sHelv As String, sHelvB As String, sTimes As String, sTimesB As String
    Dim parts() As String
    Dim I As Integer

    sHelv = "278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584"
    sHelvB = "278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584"
    sTimes = "250,333,408,500,500,833,778,180,333,333,500,564,250,333,250,278,500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,480,200,480,541"
    sTimesB = "250,333,555,500,500,1000,833,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,394,220,394,520"

    parts = Split(sHelv, ",")
    For I = 0 To 94
        m_WHelv(32 + I) = CInt(parts(I))
    Next I

    parts = Split(sHelvB, ",")
    For I = 0 To 94
        m_WHelvB(32 + I) = CInt(parts(I))
    Next I

    parts = Split(sTimes, ",")
    For I = 0 To 94
        m_WTimes(32 + I) = CInt(parts(I))
    Next I

    parts = Split(sTimesB, ",")
    For I = 0 To 94
        m_WTimesB(32 + I) = CInt(parts(I))
    Next I
End Sub

'--------------------------------------------------------------------------
' Internal: font resource management (12 base-14 fonts, always declared in
' every page's /Resources so SetFont/inline bold can switch freely)
'--------------------------------------------------------------------------

Private Sub SetFontIfNeeded(ByVal bBold As Boolean, ByVal bItalic As Boolean)
    Dim key As String
    key = m_FontFamily & "|" & bBold & "|" & bItalic & "|" & Num(m_FontSize)
    If key <> m_CurFontKey Then
        AppendRaw "/" & GetFontResName(m_FontFamily, bBold, bItalic) & " " & Num(m_FontSize) & " Tf" & vbLf
        m_CurFontKey = key
    End If
End Sub

Private Function GetFontResName(ByVal family As String, ByVal bold As Boolean, ByVal italic As Boolean) As String
    Dim baseNum As Integer
    Dim offset As Integer

    Select Case family
        Case "Times": baseNum = 4
        Case "Courier": baseNum = 8
        Case Else: baseNum = 0 ' Helvetica
    End Select

    If bold And italic Then
        offset = 3
    ElseIf bold Then
        offset = 1
    ElseIf italic Then
        offset = 2
    Else
        offset = 0
    End If

    GetFontResName = "F" & (baseNum + offset + 1)
End Function

Private Function GetBaseFontName(ByVal resName As String) As String
    Select Case resName
        Case "F1": GetBaseFontName = "Helvetica"
        Case "F2": GetBaseFontName = "Helvetica-Bold"
        Case "F3": GetBaseFontName = "Helvetica-Oblique"
        Case "F4": GetBaseFontName = "Helvetica-BoldOblique"
        Case "F5": GetBaseFontName = "Times-Roman"
        Case "F6": GetBaseFontName = "Times-Bold"
        Case "F7": GetBaseFontName = "Times-Italic"
        Case "F8": GetBaseFontName = "Times-BoldItalic"
        Case "F9": GetBaseFontName = "Courier"
        Case "F10": GetBaseFontName = "Courier-Bold"
        Case "F11": GetBaseFontName = "Courier-Oblique"
        Case "F12": GetBaseFontName = "Courier-BoldOblique"
    End Select
End Function

'--------------------------------------------------------------------------
' Internal: content-stream buffer, page buffer, and low-level binary I/O
'--------------------------------------------------------------------------

Private Sub AppendRaw(ByVal S As String)
    m_CurContent = m_CurContent & S
End Sub

Private Sub ApplyTabStops(ByVal Positions As Variant, ByVal Multiplier As Double)
    Dim I As Integer, n As Integer
    n = -1
    On Error Resume Next
    n = UBound(Positions)
    On Error GoTo 0

    If n < 0 Then
        m_TabStopCount = 0
        Exit Sub
    End If

    ReDim m_TabStops(0 To n)
    For I = 0 To n
        m_TabStops(I) = CDbl(Positions(I)) * Multiplier
    Next I
    m_TabStopCount = n + 1
End Sub

Private Function EscapePdfString(ByVal S As String) As String
    Dim r As String
    r = S
    r = Replace(r, "\", "\\")
    r = Replace(r, "(", "\(")
    r = Replace(r, ")", "\)")
    EscapePdfString = r
End Function

Private Function BuildLetterheadFooterContent() As String
    Dim savedContent As String
    Dim savedFamily As String
    Dim savedSize As Double
    Dim savedBold As Boolean
    Dim savedItalic As Boolean
    Dim savedUnderline As Boolean
    Dim savedFontKey As String
    Dim tokens() As PdfToken
    Dim nTok As Integer
    Dim lineTok() As PdfToken
    Dim lineCount As Integer
    Dim I As Integer
    Dim lineNum As Integer
    Dim spaceW As Double
    Dim wordW As Double
    Dim curWidth As Double
    Dim testWidth As Double
    Dim xStart As Double
    Dim availWidth As Double
    Dim lineH As Double
    Dim yDev As Double

    savedContent = m_CurContent
    savedFamily = m_FontFamily
    savedSize = m_FontSize
    savedBold = m_FontBold
    savedItalic = m_FontItalic
    savedUnderline = m_FontUnderline
    savedFontKey = m_CurFontKey

    m_CurContent = ""
    m_FontFamily = "Times"
    m_FontSize = 12
    m_FontBold = True
    m_FontItalic = False
    m_FontUnderline = False
    m_CurFontKey = ""

    xStart = 0
    availWidth = m_PageWidth
    lineH = LineHeight()

    TokenizeLine m_LetterheadFooterText, tokens, nTok
    If nTok = 0 Then
        BuildLetterheadFooterContent = ""
        GoTo RestoreState
    End If

    spaceW = MeasureChar(" ", m_FontFamily, False) * m_FontSize / 1000
    ReDim lineTok(0 To nTok - 1)
    lineCount = 0
    curWidth = 0
    lineNum = 0

    For I = 0 To nTok - 1
        wordW = MeasureWord(tokens(I).sWord, m_FontFamily, tokens(I).bBold) * m_FontSize / 1000
        If lineCount = 0 Then
            testWidth = wordW
        Else
            If NeedsInterWordSpaceTokens(lineTok(lineCount - 1), tokens(I)) Then
                testWidth = curWidth + spaceW + wordW
            Else
                testWidth = curWidth + wordW
            End If
        End If

        If testWidth > availWidth And lineCount > 0 Then
            yDev = m_PageHeight - ((m_LetterheadFooterY * 72) + lineNum * lineH)
            EmitLineAt lineTok, lineCount, xStart, yDev, pdfCenter, availWidth
            lineCount = 0
            curWidth = wordW
            lineNum = lineNum + 1
        Else
            curWidth = testWidth
        End If

        lineTok(lineCount) = tokens(I)
        lineCount = lineCount + 1
    Next I

    If lineCount > 0 Then
        yDev = m_PageHeight - ((m_LetterheadFooterY * 72) + lineNum * lineH)
        EmitLineAt lineTok, lineCount, xStart, yDev, pdfCenter, availWidth
    End If

    BuildLetterheadFooterContent = m_CurContent

RestoreState:
    m_CurContent = savedContent
    m_FontFamily = savedFamily
    m_FontSize = savedSize
    m_FontBold = savedBold
    m_FontItalic = savedItalic
    m_FontUnderline = savedUnderline
    m_CurFontKey = savedFontKey
End Function

Private Sub FlushCurrentPage()
    ReDim Preserve m_Pages(0 To m_PageCount)
    m_Pages(m_PageCount) = m_CurContent
    m_PageCount = m_PageCount + 1
    m_CurContent = ""
End Sub

Private Sub WriteBytes(ByVal fnum As Integer, ByRef Pos As Long, ByVal S As String)
    Dim b() As Byte
    If Len(S) = 0 Then Exit Sub
    b = StrConv(S, vbFromUnicode)
    Put #fnum, Pos, b
    Pos = Pos + (UBound(b) - LBound(b) + 1)
End Sub

' Locale-independent numeric formatting (Format$/CStr use the system's
' decimal separator, which would corrupt PDF syntax on non-US locales).
Private Function Num(ByVal v As Double) As String
    v = Int(v * 100 + 0.5 * Sgn(v)) / 100
    Num = Trim$(str$(v))
End Function
