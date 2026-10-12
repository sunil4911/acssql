VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form35 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SCAN DOCUMENTS"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14505
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   4
      Top             =   8520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   3
      Top             =   8520
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf1 
      Height          =   7935
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13996
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   8520
      Width           =   2415
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str, STR2, SCANS_ACN2 As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command_Click()

End Sub

Private Sub Command1_Click()
FR_FORM = 0
Unload Me
End Sub

Private Sub Command2_Click()
SCANS_ADD.Show (1)
End Sub

Private Sub Form_Activate()
Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
If FR_FORM = 14 Then
    Command2.Visible = False
    Command3.Visible = True
    Dim sql As String

    sql = "SELECT ACN, FORMAT([DATE],'MM-dd-yyyy'), OPC, SUBJECT, DLOC, [DEL], ID, ATYPE " & _
      "FROM ACSPDF " & _
      "WHERE [DEL] IN ('YES','APR') " & _
      "ORDER BY [DEL] DESC"

    Set rs = conn.Execute(sql)
    Set msf1.Recordset = rs
    msf1.Cols = 8
    msf1.RowHeight(-1) = 600
    With msf1
    .TextMatrix(0, 0) = "AC#"
    .TextMatrix(0, 1) = "DATE"
    .TextMatrix(0, 2) = "OPC"
    .TextMatrix(0, 3) = "SUBJECT"
    .TextMatrix(0, 4) = ""
    .TextMatrix(0, 5) = "VOID"
    
    .colWidth(0) = msf1.Width * 10 / 100
    .colWidth(1) = msf1.Width * 11 / 100
    .colWidth(2) = msf1.Width * 7 / 100
    .colWidth(3) = msf1.Width * 63 / 100
    .colWidth(4) = msf1.Width * 0 / 100
    .colWidth(5) = msf1.Width * 7 / 100
    .colWidth(6) = msf1.Width * 0 / 100
    End With

Else
    Command2.Visible = True
    Command3.Visible = False
    Set rs = conn.Execute("SELECT Format(DATE, 'MM-dd-yyyy'),OPC,SUBJECT,DLOC,DEL,ID FROM ACSPDF WHERE ACN='" & Trim(scans_acn) & "' ORDER BY DATE DESC")
    Set msf1.Recordset = rs
    msf1.Cols = 6
    msf1.RowHeight(-1) = 600
    With msf1
    .TextMatrix(0, 0) = "DATE"
    .TextMatrix(0, 1) = "OPC"
    .TextMatrix(0, 2) = "SUBJECT"
    .TextMatrix(0, 3) = ""
    .TextMatrix(0, 4) = "VOID"
    
    .colWidth(0) = msf1.Width * 13 / 100
    .colWidth(1) = msf1.Width * 8 / 100
    .colWidth(2) = msf1.Width * 68 / 100
    .colWidth(3) = msf1.Width * 0 / 100
    .colWidth(4) = msf1.Width * 8 / 100
    .colWidth(5) = msf1.Width * 0 / 100
    End With
End If
End Sub

Private Sub msF1_Click()
'strdir = Trim(scans_loc) & Trim(scans_type) & "\"
'strdir = strdir & Trim(msf1.TextMatrix(msf1.Row, 3))
'strFile = "<your filename>. pdf"
'ShellExecute 0, "OPEN", strdir, "", strdir, 1


End Sub



Private Sub msf1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Integer
    Dim c As Integer

    r = msf1.MouseRow
    c = msf1.MouseCol

    ' MouseDown fires before MSHFlexGrid updates .Row, so always use MouseRow.
    ' Using .Row here was reading the header ("DATE") and raising a type
    ' mismatch on DateValue when opening a saved PDF.
    If r < 1 Then Exit Sub

    If Button = vbLeftButton Then
        If FR_FORM = 14 Then
            If c = 0 Or c = 1 Or c = 3 Then
                scans_type = msf1.TextMatrix(r, 7)
                OpenScanDocument r, 4, 1
            End If
            If c = 5 Then
                If Trim(msf1.TextMatrix(r, 5)) = "APR" Then
                    msf1.TextMatrix(r, 5) = "YES"
                Else
                    msf1.TextMatrix(r, 5) = "APR"
                End If

                str = "update ACSPDF set "
                str = str & "[DEL]" & "='" & msf1.TextMatrix(r, 5) & "' where id=" & msf1.TextMatrix(r, 6)
                conn.Execute (str)
            End If
        Else
            If c = 0 Or c = 2 Then
                OpenScanDocument r, 3, 0
            End If
            If c = 4 Then
                If Trim(msf1.TextMatrix(r, 4)) = "YES" Then
                    msf1.TextMatrix(r, 4) = ""
                Else
                    msf1.TextMatrix(r, 4) = "YES"
                End If

                str = "update ACSPDF set "
                str = str & "[DEL]" & "='" & msf1.TextMatrix(r, 4) & "' where id=" & msf1.TextMatrix(r, 5)
                conn.Execute (str)
            End If
        End If
    End If

End Sub

Private Sub OpenScanDocument(ByVal gridRow As Integer, ByVal dlocCol As Integer, ByVal dateCol As Integer)
    Dim pdfPath As String
    pdfPath = ResolveScanPdfPath(msf1.TextMatrix(gridRow, dlocCol), msf1.TextMatrix(gridRow, dateCol))
    If Len(Trim$(pdfPath)) = 0 Then
        MsgBox "No file path is stored for this document.", vbExclamation, "SCAN DOCUMENTS"
        Exit Sub
    End If
    OpenPdfForViewing pdfPath
End Sub

' Rebuilds the on-disk path from ACSPDF.DLOC. Newer rows already store a
' full absolute path; older rows store a filename (or a path relative to
' scans_loc) and need the scans_loc / scans_type prefix. Try the likely
' locations and keep the date cutoff only as a last resort, and never call
' DateValue unless IsDate succeeds.
Private Function ResolveScanPdfPath(ByVal dloc As String, ByVal docDate As String) As String
    Dim p As String
    dloc = Trim$(dloc)
    If Len(dloc) = 0 Then Exit Function

    If IsAbsolutePdfPath(dloc) Then
        ResolveScanPdfPath = dloc
        Exit Function
    End If

    p = Trim$(scans_loc) & dloc
    If PdfFileExists(p) Then
        ResolveScanPdfPath = p
        Exit Function
    End If

    p = Trim$(scans_loc) & Trim$(scans_type) & "\" & dloc
    If PdfFileExists(p) Then
        ResolveScanPdfPath = p
        Exit Function
    End If

    If IsDate(docDate) Then
        If DateValue(docDate) < DateValue("03/01/2026") Then
            ResolveScanPdfPath = Trim$(scans_loc) & Trim$(scans_type) & "\" & dloc
        Else
            ResolveScanPdfPath = Trim$(scans_loc) & dloc
        End If
    Else
        ResolveScanPdfPath = Trim$(scans_loc) & dloc
    End If
End Function

' ACSPDF.DLOC has been saved in a few different formats over time - some
' older rows are just a bare filename or a path relative to scans_loc, but
' the current save code (save_Pleadings in ACSSQL.bas, GenerateLet54PDF in
' let26.frm) always writes the full absolute path (e.g. "u:\acs\pdf\2023\
' 202345\...pdf"). Re-prefixing scans_loc onto an already-absolute path
' produces a bogus doubled path (e.g. "u:\acs\pdf\u:\acs\pdf\...") that
' can't be opened, which is what was causing the error when opening saved
' PDFs from this form. Detect an absolute path (drive letter or UNC) so we
' only add the scans_loc/scans_type prefix for the older, non-absolute DLOC
' values that actually still need it.
Private Function IsAbsolutePdfPath(ByVal p As String) As Boolean
    p = Trim$(p)
    If Len(p) >= 2 Then
        If mid$(p, 2, 1) = ":" Then
            IsAbsolutePdfPath = True
            Exit Function
        End If
        If Left$(p, 2) = "\\" Then
            IsAbsolutePdfPath = True
            Exit Function
        End If
    End If
End Function

