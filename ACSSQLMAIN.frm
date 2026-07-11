VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACSMAIN"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   6930
   Begin VB.CommandButton Command9 
      Caption         =   "pleadings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3360
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command7"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   7680
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "INFO SUBPOENA (clsPDF)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   10
      Top             =   8640
      Width           =   4425
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TABLE UPDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1080
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   4425
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1200
      Top             =   9000
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=rss2024"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "rss2024"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   4
      Top             =   6960
      Width           =   4425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MISCELLANEOUS TASKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   3
      Top             =   4620
      Width           =   4425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MANAGEMENT REPORTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   2
      Top             =   3465
      Width           =   4425
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "MONTHLY TASKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   1
      Top             =   2280
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DAILY TASKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1050
      TabIndex        =   0
      Top             =   1155
      Width           =   4425
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   6600
      TabIndex        =   5
      Top             =   5160
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub
Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub
Private Sub Command3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub
Private Sub Command4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub
Private Sub Command5_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub
Private Sub Command6_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    End
End If
End Sub

Private Sub Command4_Click()
MISC_TASKS.Show (1)
End Sub
Private Sub Command2_Click()
Form13.Show
End Sub

Private Sub Command3_Click()
Form14.Show
End Sub

Private Sub Command5_Click()
End
End Sub



Private Sub Command6_Click()
'Form36.Show (1)
End Sub


Private Sub Command7_Click()
Dim dr As DataReport
Dim I As Integer
I = 1
'Set dr = "stdlh"
'STDLETTERHEAD STDLH, "Section4", 1
STDLETTERHEAD DataReport18, "Section1", 1
STDLETTERHEAD DataReport18, "Section7", 24
'STDLETTERHEAD DataReport18, "Section6", 48
DataReport18.Show (1)
'STDLH.Show (1)
'DataEnvironment1.Command15 MOPC
'DataReport16.Show (1)
End Sub

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
PDF.PrintTabbedLine "Item" & vbTab & "Qty" & vbTab & "Price", Array(pdfLeft, pdfRight, pdfRight)
PDF.SetFont "Helvetica", 11, False
PDF.PrintTabbedLine "Widget" & vbTab & "3" & vbTab & "**$29.97**", Array(pdfLeft, pdfRight, pdfRight)
PDF.PrintTabbedLine "Gadget" & vbTab & "1" & vbTab & "**$14.99**", Array(pdfLeft, pdfRight, pdfRight)

PDF.PrintText ""
PDF.PrintText "Please remit payment within __10 days__ of this notice.", pdfJustify

PDF.SavePDF


 
    'Dim PDF As New clsPDF
    
 'PDF.CreatePDF "C:\TEMP\TEST.PDF"
 'PDF.SetMarginsPt 36, 36, 36, 36
 'PDF.SetFont "F1", 12
 
 'PDF.PrintParagraph _
'"This is a very long paragraph used to test the page margins. The purpose is to verify that the text wraps correctly between the left and right margins. This paragraph contains enough words so that several lines are created and we can see whether the right edge is staying at the correct half inch margin. The document generator should handle letters, agreements, and legal pleadings correctly."

'PDF.SavePDF
    MsgBox "DONE"
    

End Sub


' Runs the clsPDF-based Information Subpoena package (cover letter +
' subpoena + accompanying questionnaire, all sharing the Letterhead_clsPDF
' template - see LET10_InfoSubpoena_clsPDF.bas and Letterhead_clsPDF.bas).
'
' Uses the current case's MCOURT/MCOUNTY/PL_P/PL_D/PL_C globals when a case
' is loaded (non-blank); otherwise falls back to sample data so the button
' always produces a valid demo document, the same way Command8's Invoice
' demo works. IndexNo, DebtorAddress, judgment date/amounts, and the
' attorney name don't have an unambiguous global source yet (the legacy
' let10.frm pulled Index No from Form12.Text1(11) and the debtor address via
' an SSN lookup against AAP) - wire those in here once you tell me where
' they should come from for a live case.
Private Sub Command10_Click()
    Dim OutPath As String
    Dim CourtNameV As String, CountyLineV As String
    Dim PlaintiffV As String, DefendantV As String, CoDefendantV As String

    CourtNameV = IIf(Trim(MCOURT) <> "", MCOURT, "Oswego Supreme Court")
    CountyLineV = IIf(Trim(MCOUNTY) <> "", MCOUNTY, "County of Oswego State of New York")
    PlaintiffV = IIf(Trim(PL_P) <> "", PL_P, "Oswego County Federal Credit Union")
    DefendantV = IIf(Trim(PL_D) <> "", PL_D, "Trevor S Brooks")
    CoDefendantV = Trim(PL_C)

    OutPath = "C:\temp\LET10_Package.pdf"

    PrintSubpoenaWithCoverLetter _
        PdfPath:=OutPath, _
        RecipientLines:=DefendantV & vbNewLine & "40 Hughes Rd.  Mexico, NY 13114", _
        DateStr:=Format$(Date, "mm/dd/yy"), _
        Salutation:="Dear Sir or Madam:", _
        ReferenceLine:="Index No. EFC-2025-0038", _
        CoverLetterBody:= _
            "Enclosed please find an Information Subpoena in the above-referenced matter. " & _
            "As set forth in the subpoena, you are required to answer the accompanying " & _
            "questionnaire under oath and return it, together with the original questions, " & _
            "within seven (7) days of your receipt of this letter.", _
        AttorneyName:="Amanda C Shaw", _
        CourtName:=CourtNameV, _
        CountyLine:=CountyLineV, _
        IndexNo:="EFC-2025-0038", _
        PlaintiffName:=PlaintiffV, _
        DefendantName:=DefendantV, _
        DebtorAddress:="40 Hughes Rd.  Mexico, NY 13114", _
        JudgmentDateStr:="05/13/25", _
        AmountOriginal:="$9,423.88", _
        AmountDue:="$9,778.40", _
        CoDefendantName:=CoDefendantV

    ShellExecute 0, "OPEN", OutPath, "", "", 1
End Sub

Private Sub Command9_Click()
MLET = 10
PLEADINGS_DEF.Show (1)
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim str As String
'STR = "2026-05-19"
'Command5.Caption = Format$(STR, "mmmm dd, yyyy")
'================COMPANY INFORMATION====================

Set rs = conn.Execute("select * from ACS")
If rs.EOF = True And rs.BOF = True Then
Else
   For I = 1 To rs.Fields.Count
      MACS(I) = ""
   Next
   For I = 0 To rs.Fields.Count - 1
      MACS(I + 1) = Trim(rs.Fields(I))
   Next
End If
'==================================================
If Trim(MOPC) = "SGT" Then Command9.Visible = True

Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
