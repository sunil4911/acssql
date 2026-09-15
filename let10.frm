VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form910 
   BackColor       =   &H000000FF&
   Caption         =   "INFO SUB"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   7395
   ScaleWidth      =   14355
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   14
      Top             =   4680
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3000
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM-dd-yyyy"
      Format          =   146341891
      CurrentDate     =   45752
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   11
      Top             =   2040
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9720
      TabIndex        =   10
      Text            =   "Amanda C Shaw"
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   9720
      TabIndex        =   12
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11040
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   6120
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DT2 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM-dd-yyyy"
      Format          =   146341891
      CurrentDate     =   45752
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   495
      Left            =   11160
      TabIndex        =   16
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM-dd-yyyy"
      Format          =   146341891
      CurrentDate     =   45752
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Caption         =   "DELIVERY DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11160
      TabIndex        =   17
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      Caption         =   "SALUTATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "TICKER FOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "EMPLOYEE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9720
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "REVIEW DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "PRINT DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs, DS As New ADODB.Recordset
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub Combo1_Click()
Dim rs As New ADODB.Recordset
Set rs = conn.Execute("select code,concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name,FULL_PART from employee where concat(trim(fn),' ',trim(mi),' ',trim(ln))='" & Trim(Combo1.Text) & "'")
EMP2 = rs.Fields(0)
If rs.Fields(2) = "A" Then
   EMP_TYPE = "A"
Else
   EMP_TYPE = "E"
End If

End Sub

Private Sub Combo2_Click()
Dim rs As New ADODB.Recordset
Set rs = conn.Execute("select code,concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name from employee where concat(trim(fn),' ',trim(mi),' ',trim(ln)) like '" & Trim(Combo2.Text) & "%'")
EMP = rs.Fields(0)

End Sub

Private Sub Command1_Click()
Dim I, J, T_str, STR2, str3, STR4, TOP_HEIGHT As Integer
Dim str, saluta As String
Dim address() As String
Dim lineCount As Long

Dim re(10) As String
Dim signor(10) As String
Dim encs(10) As String
Dim dr As DataReport
Dim rs, DS As New ADODB.Recordset
Dim ddate As Date
Dim filesize, mchg, mint, MPMT, MRATE, MDAYS As Double

Set rs = conn.Execute("select concat(trim(fn),' ',trim(mi),' ',trim(ln)),code,TITLE as name from employee where concat(trim(fn),' ',trim(mi),' ',trim(ln)) ='" & Trim(Combo1.Text) & "'")
If (rs.EOF = True And rs.BOF = True) Or IsNull(rs.Fields(2)) = True Or Trim(rs.Fields(2)) = "" Then
    Label6.Caption = "Account Representative"
Else
    Label6.Caption = rs.Fields(2)
End If
signor(1) = Trim(Combo1.Text)
signor(2) = Trim(Label6.Caption)

If DATA_VALIDATE(Trim(Combo2.Text), "concat(trim(fn),' ',trim(mi),' ',trim(ln))", "EMPLOYEE") = 0 Then
    MESSAGE_TEXT = "INVALID TICKER"
    MSG_TEXT.Show (1)
    MESSAGE_TEXT = ""
    Exit Sub
Else
   EMP = rs.Fields(1)
End If
If DATA_VALIDATE(Trim(Combo1.Text), "concat(trim(fn),' ',trim(mi),' ',trim(ln))", "EMPLOYEE") = 0 Then
    MESSAGE_TEXT = "INVALID EMPLOYEE"
    MSG_TEXT.Show (1)
    MESSAGE_TEXT = ""
    Exit Sub
End If

' Case 12 uses clsPDF directly; starting Acrobat here causes
' "Component Request Pending" if Adobe Reader is busy or open.
If MLET <> 12 Then
    Dim ap As Acrobat.AcroApp
    Dim dr1 As Acrobat.AcroPDDoc
    Dim dr2 As Acrobat.AcroPDDoc
    Set ap = New Acrobat.AcroApp
    Set dr1 = New Acrobat.AcroPDDoc
    Set dr2 = New Acrobat.AcroPDDoc

    Dim prt As Printer
    For Each prt In Printers
        If prt.DeviceName = "PDF995" Then
            Set Printer = prt
            Exit For
        End If
    Next
End If

address = Split(Text1(0).Text, vbCrLf)

Dim TEMP As Variant
TEMP = Split(PL_P, vbNewLine)
lineCount = UBound(TEMP) - LBound(TEMP) + 1
If lineCount > 1 Then
    For I = 1 To lineCount
        If Len(Trim$(TEMP(I - 1))) > 0 Then
            re(I) = Trim(TEMP(I - 1))
            J = I + 1
        End If
    Next I
Else
    J = 2
    re(1) = Trim(PL_P)
End If

Dim TEMP2 As Variant
TEMP2 = Split(PL_D, vbNewLine)
lineCount = UBound(TEMP2) - LBound(TEMP2) + 1
If lineCount > 1 Then
    For I = 1 To lineCount
        If Len(Trim$(TEMP2(I - 1))) > 0 Then
            If I = 1 Then
                re(I + J) = "vs: " & Trim(TEMP2(I - 1))
            Else
                re(I + J) = Trim(TEMP2(I - 1))
            End If
        End If
    Next I
Else
    re(J) = "vs: " & Trim(PL_D)
End If

Select Case MLET
    Case 10, 12, 13:
                
        Dim PDF As New clsPDF
        Dim par(10) As String
        Dim OutPdf As String
        'Dim I As String
        If MLET = 10 Then
            par(1) = "Enclosed is an Information Subpoena and a questionnaire to be completed by you and returned to"
            par(1) = par(1) & "this office within seven (7) days of receipt of same. Please sign where indicated on the form and"
            par(1) = par(1) & "have your signature notarized. You may keep one copy of the questionnaire and return the other"
            par(1) = par(1) & "copy to this office in the envelope provided."
            par(2) = "Please be advised that failure to comply with this subpoena may be punishable as contempt of court."
            encs(1) = "/" & Trim(UCase(MOPC))
            encs(2) = "encs."
        End If
        If MLET = 12 Then
            If Trim(Form12.Text1(7).Text) <> "" Then
                par(1) = "We have not received the answers to the Information Subpoena that your client received on "
                par(1) = par(1) & DT3 & ". Kindly have your client complete the Questionnaire and return it to this office within (5) days."
            Else
                par(1) = "We have not received the answers to the Information Subpoena which "
                par(1) = par(1) & "was delivered to you on " & DT3 & "."
                par(2) = "It is required by law that the questionnaire be completed by you"
                par(2) = par(2) & " and returned to this office within seven (7) days of receipt of"
                par(2) = par(2) & " same. If the questionnaire is not received by this office within"
                par(2) = par(2) & " the next five (5) days, we may make a'"
                par(2) = par(2) & " motion in Supreme Court to have you held in contempt of court. If"
                par(2) = par(2) & " you fail to appear in Court at the time of the motion, it could"
                par(2) = par(2) & " result in your immediate arrest and imprisonment for Contempt of Court."
            End If
            encs(1) = "/" & Trim(UCase(MOPC))
            encs(2) = "encs."
        End If
        If MLET = 13 Then
            par(1) = "Our next step will be to ask the court to hold you in contempt because"
            par(1) = par(1) & " you have failed to answer the Information Subpoena which was delivered"
            par(1) = par(1) & " to you on " & DT3 & ". We strongly urge you to consult an attorney about"
            par(1) = par(1) & "' the penalties that may be imposed upon a finding of contempt."
            par(2) = "The violation is the failure to answer the subpoena. You must answer the subpeona."
            par(2) = par(2) & "' Please fill out the questionnaire in full and return it to our office"
            par(2) = par(2) & " or we may have no alternative but to ask for a finding of contempt."
            par(3) = "If you have any questions, please call our office and we will be happy"
            par(3) = par(3) & " to talk with you about this matter. You may find that the problem is not as big as you think."
        End If
        
        
        OutPdf = "C:\temp\LET10.pdf"
        PDF.CreatePDF OutPdf
        PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
        PDF.SetLetterheadFooter LetterheadFooterText()            ' bottom disclosure (variable text)
        PDF.SetMargins 0.5, 1.5, 2.75, 1
        'MsgBox FormatLongDate(DT1)
        PDF.SetFont "Times New Roman", 12, False
        PDF.SetStartX 3.25
        PDF.PrintText FormatLongDate(DT1)
        PDF.PrintText ""
        PDF.PrintText ""
        For I = 0 To UBound(address)
            If Len(Trim$(address(I))) > 0 Then PDF.PrintText Trim$(address(I)), pdfLeft
        Next I
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.SetStartX 1
        PDF.PrintText "Re: " & re(1), pdfLeft
        For I = 2 To 10
            If Len(Trim$(re(I))) > 0 Then
                PDF.SetStartX 1.25
                PDF.PrintText Trim$(re(I)), pdfLeft
            End If
        Next I
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.PrintText Trim(Text1(1).Text), pdfLeft
        PDF.PrintText ""
        For I = 1 To 10
            If Len(Trim$(par(I))) > 0 Then
                PDF.SetStartX 0.25
                PDF.PrintText Trim$(par(I)), pdfJustify
                PDF.PrintText ""
            End If
        Next I
        PDF.SetStartX 3.25
        PDF.PrintText "Very truly yours,", pdfLeft
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.PrintText ""
        For I = 1 To 10
            PDF.SetStartX 3.25
            If Len(Trim$(signor(I))) > 0 Then PDF.PrintText Trim$(signor(I)), pdfLeft
        Next I
        
        PDF.PrintText ""
        For I = 1 To 10
            If Len(Trim$(encs(I))) > 0 Then PDF.PrintText Trim$(encs(I)), pdfLeft
        Next I
        If MLET = 10 Then
            Dim CourtName, CountyLine, IndexNo, PlaintiffName, DefendantName, DebtorAddress As String
            Dim JudgmentDateStr, AmountOriginal, AmountDue As String
            Dim SignerName, FirmName, FirmTagline1, FirmAddressLine1, FirmAddressLine2 As String
            Dim CoDefendantName, DateStr As String
            Dim IndentIn As Double
            CoDefendantName = ""
            DateStr = DT1
            CourtName = "Oswego Supreme Court"
            CountyLine = "County of Oswego State of New York"
            IndexNo = "EFC-2025-0038"
            PlaintiffName = "Oswego County Federal Credit Union"
            DefendantName = "Trevor S Brooks"
            'Dim CoDefendantName As String
            DebtorAddress = "40 Hughes Rd.  Mexico, NY 13114"
            JudgmentDateStr = "05/13/25"
            AmountOriginal = "$9,423.88"
            AmountDue = "$9,778.40"
            'ByVal CourtName As String, _
            'ByVal CountyLine As String, _
            'ByVal IndexNo As String, _
            'ByVal PlaintiffName As String, _
            'ByVal DefendantName As String, _
            'ByVal DebtorAddress As String, _
            'ByVal JudgmentDateStr As String, _
            'ByVal AmountOriginal As String, _
            'ByVal AmountDue As String, _
            'ByVal DateStr As String, _
            'ByVal AttorneyName As String, _
            'ByVal FirmName As String, _
            'ByVal FirmAddressLine1 As String, _
            'ByVal FirmAddressLine2 As String, _
            'ByVal CoDefendantName As String)
        
            'Dim I As Integer
            PDF.NewPage
            PDF.SetMargins 0.65, 0.25, 0.65, 0.5
            PDF.SetFont "Times", 12, False
 
            DrawCaptionBox PDF, CourtName, CountyLine, IndexNo, PlaintiffName, DefendantName, _
                CoDefendantName, "INFORMATION SUBPOENA"
        
            ' --- Bold centered heading ----------------------------------------------
            PDF.SetFont "Times New Roman", 12, True
            PDF.PrintText "THE PEOPLE OF THE STATE OF NEW YORK", pdfCenter
            PDF.SetFont "Times New Roman", 12, False
            PDF.PrintText ""
        
            ' --- Recipient / body -----------------------------------------------------
            PDF.PrintText "TO:  " & DefendantName & ", the person to be examined", pdfLeft
            PDF.PrintText "ADDRESS:  " & DebtorAddress, pdfLeft
            PDF.PrintText ""
            PDF.PrintText "GREETINGS:", pdfLeft
            PDF.PrintText ""
        
            PDF.PrintText "     WHEREAS, in an action in the " & CourtName & " of " & CountyLine & _
                " between " & PlaintiffName & " , as plaintiff, and " & DefendantName & ", as defendant, who " & _
                "are all the parties named in said action, a judgment was entered on " & JudgmentDateStr & _
                " in favor of " & PlaintiffName & " , judgment creditor, and against " & DefendantName & _
                ", judgment debtor in the amount of " & AmountOriginal & " of which " & AmountDue & _
                " remains due and unpaid, which includes all costs and interest to date; and", pdfJustify
            PDF.PrintText ""
        
            PDF.PrintText "     NOW, THEREFORE WE COMMAND YOU, that you answer in writing under oath, separately and " & _
                "fully, each question in the questionnaire accompanying this subpoena, each answer referring to the " & _
                "question to which it responds; and that you return the answers together with the original of the questions " & _
                "within seven (7) days after your receipt of the questions and this subpoena.", pdfJustify
            PDF.PrintText ""
        
            PDF.PrintText "     TAKE NOTICE that false swearing or failure to comply with this subpoena is punishable as " & _
                "contempt of court.", pdfJustify
            PDF.PrintText ""
            PDF.PrintText ""
        
            PDF.PrintText "I HEREBY CERTIFY THAT THIS INFORMATION SUBPOENA COMPLIES WITH RULE 5224 OF " & _
                "THE CIVIL PRACTICES LAW AND THAT I HAVE A REASONABLE BELIEF THAT THE PARTY " & _
                "RECEIVING THIS SUBPOENA HAS IN THEIR POSSESSION INFORMATION ABOUT THE " & _
                "DEBTOR THAT WILL ASSIST THE CREDITOR IN COLLECTING THE JUDGMENT.", pdfJustify
            PDF.PrintText ""
            PDF.PrintText ""
            PDF.PrintText ""
        
            ' --- Signature block (shared template piece) --------------------------------
            PDF.PrintText "DATE: " & DateStr, pdfLeft
        
            PrintFirmSignatureBlock PDF, _
                SignerName = signor(1), _
                FirmName = MACS(1), _
                FirmTagline1 = "Attorneys & Counselors at Law", _
                FirmAddressLine1 = MACS(3), _
                FirmAddressLine2 = MACS(5), _
                IndentIn = 4.5
        
            ' --- Bottom disclosure block (shared template piece; approximate position
            ' near page bottom - clsPDF flows sequentially with no absolute-Y
            ' positioning, so this pads down with blank lines rather than anchoring
            ' precisely to the bottom) --------------------------------------------------
            For I = 1 To 2
                PDF.PrintText ""
            Next I
        
            PrintFDCPADisclosure PDF, FirmName:=FirmName
        End If
        PDF.SavePDF
        OpenPdfForViewing PDF.FilePath
End Select
'PLEADING_UPDATE
Unload Me
'DataReport14.PrintReport
'ShellExecute 0, "OPEN", "t:\stmt.pdf", "", "t:\stmt.pdf", 1

End Sub

Private Sub Command2_Click()
MsgBox EMP
End Sub

Private Sub DT1_Change()
Dim rs As New ADODB.Recordset
Set rs = conn.Execute("Select REV_DAYS from SAF where code=26")
If IsNull(rs.Fields(0)) = False Then
    DT2.Value = Date + Val(rs.Fields(0))
Else
    DT2.Value = Date + 14
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim I As Integer
Dim str As String
'Dim EMP_TYPE As String
Me.KeyPreview = True
'EMP = MOPC
EMP = Form12.Text1(55).Text
Text1(4).Text = Form12.Text1(57)
EMP2 = "ACS"
'EMP_TYPE = "A"
Set rs = conn.Execute("Select REV_DAYS, name,signed_by from SAF where code=" & MLET)
DT1.Value = Date
If IsNull(rs.Fields(0)) = False Then
    DT2.Value = Date + Val(rs.Fields(0))
Else
    DT2.Value = Date + 14
End If
Me.Caption = Trim(rs.Fields(1))

If IsNull(rs.Fields(2)) = False And Trim(rs.Fields(2)) = "A" Then
    EMP_TYPE = "A"
    Set rs = conn.Execute("select concat(trim(fn),' ',trim(mi),' ',trim(ln)),code as name from employee where upper(full_part)='A' AND (last_day is null or last_day='0000-00-00') order by name")
    Do While rs.EOF = False
        Combo1.AddItem Trim(rs.Fields(0))
        rs.MoveNext
    Loop
Else
    EMP_TYPE = "C"
    Set rs = conn.Execute("Select concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name,title,code from employee where (last_day is null or last_day='0000-00-00') order by name")
    Do While rs.EOF = False
        Combo1.AddItem Trim(rs.Fields(0))
        rs.MoveNext
    Loop
End If

Set rs = conn.Execute("Select concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name,title,code,CALENDAR from employee where (last_day is null or last_day='0000-00-00') and (upper(calendar)<>'N' or upper(calendar) is not null)  order by name")
Do While rs.EOF = False
    Combo2.AddItem Trim(rs.Fields(0))
    rs.MoveNext
Loop
Combo1.Text = EMP_NAME
If EMP_TYPE = "A" Then Combo1.Text = MACS(8)
Combo2.Text = EMP_NAME
DT2.MinDate = Date + 1
Set rs = conn.Execute("SELECT CALENDAR FROM EMPLOYEE WHERE CODE='" & Trim(MOPC) & "'")
If IsNull(rs.Fields(0).Value) Or rs.Fields(0).Value <> "Y" Then
'If rs.Fields(0) <> "Y" Then
    Set rs = conn.Execute("SELECT concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name,CALENDAR FROM EMPLOYEE WHERE CODE='" & Trim(Form12.Text1(55).Text) & "'")
    Combo2.Text = rs.Fields(0)
End If
'Combo1.SelText
'C:\ACSSQL\pdf
If Trim(Form12.Text1(7).Text) = "" Then
    str = ""
    Set rs = conn.Execute("SELECT FN,MI,LN,bname1,bname2,bname3,ADR1,Adr2,CITY,STATE,ZIP,saluta FROM AAP WHERE SS='" & Trim(SSN) & "'")
    If rs.Fields(0) <> "" Then
       str = Trim(rs.Fields(0)) & " " & Trim(rs.Fields(1)) & " " & Trim(rs.Fields(2))
    Else
       str = Trim(rs.Fields(3))
    End If
    str = str & vbNewLine & Trim(rs.Fields(6))
    If rs.Fields(7) <> "" Then str = str & vbNewLine & Trim(rs.Fields(7))
    str = str & vbNewLine & Trim(rs.Fields(8)) & " " & Trim(rs.Fields(9)) & " " & Trim(rs.Fields(10))
    Text1(0).Text = str
    If IsNull(rs.Fields(0)) = False Then
        Text1(1).Text = "Dear " & Trim(rs.Fields("SALUTA")) & " " & Trim(rs.Fields(2)) & ":"
    Else
        Text1(1).Text = "Dear Sir/Madam" & ":"
    End If
Else
    str = ""
    I = InStr(Trim(Form12.Text1(7).Text), "-")
    If I > 1 Then
        str = Left(Trim(Form12.Text1(7).Text), I - 1)
    Else
        str = Trim(Form12.Text1(7).Text)
    End If
    Set rs = conn.Execute("SELECT NAME,ADDRESS,ADDR2,CITY,STATE,ZIP FROM LMF where code= '" & Trim(str) & "'")
    str = Trim(rs.Fields(0))
    str = str & vbNewLine & Trim(rs.Fields(1))
    If rs.Fields(2) <> "" Then str = str & vbNewLine & Trim(rs.Fields(2))
    str = str & vbNewLine & Trim(rs.Fields(3)) & " " & Trim(rs.Fields(4)) & " " & Trim(rs.Fields(5))
    Text1(0).Text = str
    Text1(1).Text = "Dear Sir/Madam" & ":"
End If
End Sub
Private Sub Form_Activate()
Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
'If GetAsyncKeyState(VK_ESCAPE) = 1 Then Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub DrawCaptionBox(ByRef PDF As clsPDF, _
    ByVal CourtName As String, _
    ByVal CountyLine As String, _
    ByVal IndexNo As String, _
    ByVal PlaintiffName As String, _
    ByVal DefendantName As String, _
    ByVal CoDefendantName As String, _
    ByVal HeadingLine1 As String, _
    Optional ByVal HeadingLine2 As String = "")

    Dim CaptionRule As String
    CaptionRule = String$(50, Chr$(175))

    PDF.PrintText CourtName, pdfLeft
    PDF.PrintText CountyLine, pdfLeft
    PDF.PrintText CaptionRule, pdfLeft

    PDF.SetTabStops 0, 3.8
    PDF.PrintTabbedLine PlaintiffName & vbTab & "Index  " & IndexNo, _
        Array(pdfLeft, pdfLeft)

    PDF.SetStartX 2
    PDF.PrintText "Plaintiff (s)", pdfLeft

    PDF.SetStartX 3.8
    PDF.PrintText HeadingLine1, pdfLeft
    If Len(HeadingLine2) > 0 Then
        PDF.SetStartX 3.8
        PDF.PrintText HeadingLine2, pdfLeft
    End If

    PDF.SetStartX 2
    PDF.PrintText "-vs-", pdfLeft

    PDF.PrintText DefendantName, pdfLeft
    If Len(CoDefendantName) > 0 Then PDF.PrintText CoDefendantName, pdfLeft

    PDF.SetStartX 2
    PDF.PrintText "Defendant (s)", pdfLeft
    PDF.PrintText CaptionRule, pdfLeft
    PDF.PrintText ""
End Sub
Private Sub PrintFirmSignatureBlock(ByRef PDF As clsPDF, _
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
