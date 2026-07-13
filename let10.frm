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
      Format          =   147718147
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
      Format          =   147718147
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
      Format          =   147718147
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



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
Dim I, T_str, STR2, str3, STR4, TOP_HEIGHT As Integer
Dim str, saluta As String
Dim address() As String
Dim re(10) As String
Dim signor(10) As String
Dim encs(10) As String
Dim dr As DataReport
Dim rs, DS As New ADODB.Recordset
Dim ap As Acrobat.AcroApp
Dim dr1 As Acrobat.AcroPDDoc
Dim dr2 As Acrobat.AcroPDDoc
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

address = Split(Text1(0).Text, vbCrLf)
re(1) = Trim(CRE_ADR(1))
Select Case MLET
    Case 10:
        T_str = "TO:" & Space(3) & vbTab
            str = ""
            For I = 1 To 5
                str = str & vbNewLine
            Next I
            PRINT_DATE = DT1
            str = str & Chr(160) & Space(70) & MonthName(Month(DT1)) & " " & Day(DT1) & ", " & Year(DT1) & vbNewLine & vbNewLine
            If Len(Trim(Text1(0).Text)) = 0 Then
                For I = 0 To 10
                    If Trim(rs.Fields(I)) <> "" Then
                        str = str & rs.Fields(I) & " "
                        T_str = T_str & rs.Fields(I) & " "
                        mid_str = mid_str & rs.Fields(I) & " "
                        If I > 1 And I < 7 Then
                            str = str & vbNewLine
                            T_str = T_str & vbNewLine & Chr(160) & Space(8) & vbTab
                            mid_str = mid_str & vbNewLine
                        End If
                    End If
                Next I
            Else
                str = str & Trim(Text1(0).Text)
            End If
            str = str & vbNewLine & vbNewLine & vbNewLine & Chr(160) & Space(10) & "Re: " & Trim(PL_P) + vbNewLine
            str = str & Chr(160) & Space(10) & "      vs: " & Trim(PL_D) + vbNewLine
            
            REV_REM = Text1(4).Text
            PRINT_DATE = DT1
            REV_DATE = DT2
            str = str & vbNewLine & vbNewLine
            If Len(Trim(Text1(0).Text)) = 0 Then
                If IsNull(rs.Fields(0)) = False Then
                    str = str & "Dear " & Trim(rs.Fields("SALUTA")) & " " & Trim(rs.Fields(2)) & ";" & vbNewLine & vbNewLine
                Else
                    str = str & "Dear Sir/Madam" & ":" & vbNewLine & vbNewLine
                End If
            Else
                str = str & Trim(Text1(1).Text) & vbNewLine & vbNewLine
            End If
        Set rs = conn.Execute("Select DES from PLEADINGS where Pcode=" & MLET)
        str = str & Trim(rs.Fields(0)) & vbNewLine & vbNewLine
        str = str & Chr(160) & Space(70) & "Very truly yours," & vbNewLine & vbNewLine & vbNewLine
        str = str & String$(70, Chr$(160)) & "____________________" & vbCrLf
        str = str & Chr(160) & Space(70) & Combo1.Text & vbNewLine
        str = str & Chr(160) & Space(70) & Label6.Caption & vbNewLine
        str = str & "/" & Trim(MOPC) & vbNewLine
        str = str & "enc."
        
        REV_DATE = Format(DT2, "YYYY-MM-YY")
        REV_REM = Text1(4).Text
        Set DS = conn.Execute("select code,concat(trim(fn),' ',trim(mi),' ',trim(ln)) as name from employee where concat(trim(fn),' ',trim(mi),' ',trim(ln)) like '" & Trim(Combo2.Text) & "%'")
        EMP = DS.Fields(0)
        
        '====================LETTER===============================================
        DATAREPORT930.TopMargin = 100
        lhbody = str
        STDLETTERHEAD DATAREPORT930, "section10", 301
        
        '=========================================================================
        
        'TOP SECTION ================================================================
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL7").Caption = "*" & Trim(MCASE) & "*"
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL35").Caption = "QUESTIONNAIRE IN CONNECTION WITH INFORMATION SUBPOENA"
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL35").Top = 1500
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL36").Caption = "Index No " & Trim(Form12.Text1(11).Text)
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL36").Top = 800
        
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL999").Caption = "*" & Trim(MCASE) & "*"
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL4").Caption = "INFORMATION SUBPOENA"
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL4").Top = 1550
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL3").Caption = "Index No " & Trim(Form12.Text1(11).Text)
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL3").Top = 800
        
        
        
        
        str = ""
        str = vbNewLine & vbNewLine
        str = MCOURT & vbNewLine & UCase(MCOUNTY) & vbNewLine
        str = str & String(50, Chr(175)) & vbNewLine
        str = str & UCase(PL_P) & vbNewLine
        str = str & Chr(160) & Space(40) & "Plaintiff(s)" & vbNewLine
        str = str & Chr(160) & Space(40) & "-vs-" & vbNewLine
        If Trim(PL_C) <> "" Then
           str = str & UCase(PL_D) & vbNewLine & UCase(PL_C) & vbNewLine
        Else
           str = str & UCase(PL_D) & vbNewLine
        End If
        str = str & Chr(160) & Space(40) & "Defendant(s)" & vbNewLine
        str = str & String(50, Chr(175))
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL2").Top = 288
                DATAREPORT930.Sections("Section4").Controls.Item("LABEL2").Caption = str
                TOP_HEIGHT = 100 + (DATAREPORT930.Sections("Section4").Controls.Item("LABEL2").Height)
                DATAREPORT930.Sections("Section4").Height = TOP_HEIGHT
        
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL10").Top = 288
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL10").Caption = str
        TOP_HEIGHT = 100 + (DATAREPORT930.Sections("Section6").Controls.Item("LABEL10").Height)
        DATAREPORT930.Sections("Section6").Height = TOP_HEIGHT
        
        'DETAIL SECTION =============================================================
        
        str = ""
        str = str & vbNewLine & String$(50, Chr$(160)) & "THE PEOPLE OF THE STATE OF NEW YORK" & vbCrLf
        str = str & vbNewLine
        str = str & "TO: " & Trim(PL_D) & ", the person to be examined." & vbNewLine
        'str = str & "ADDRESS: " & ADDRESS & vbNewLine & vbNewLine
        str = str & "GREETINGS:" & vbNewLine & vbNewLine
        str = str & "WHEREAS, in an action in the Oswego Supreme Court of County of Oswego State of New York"
        str = str & "between Oswego County Federal Credit Union , as plaintiff, and Trevor S Brooks, as defendant, who"
        str = str & "are all the parties named in said action, a judgment was entered on 05/13/25 in favor of Oswego County"
        str = str & "Federal Credit Union , judgment creditor, and against Trevor S Brooks, judgment debtor in the amount"
        str = str & "of $9,423.88 of which $9,778.40 remains due and unpaid, which includes all costs and interest to date;" & vbNewLine
        str = str & "and" & vbNewLine & vbNewLine
        Set rs = conn.Execute("Select DES2 from PLEADINGS where Pcode=" & MLET)
        str = str & Trim(rs.Fields(0)) & vbNewLine & vbNewLine & vbNewLine
        str = str & String$(70, Chr$(160)) & "______________________________" & vbCrLf
        str = str & Chr(160) & Space(70) & Combo1.Text
        For I = 1 To 6
            If Len(MACS(I)) > 5 Then
                str = str & vbNewLine & Chr(160) & Space(70) & Trim(MACS(I))
            End If
        Next
        str = str & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        DATAREPORT930.Sections("Section4").Controls.Item("LABEL1").Caption = str
        DATAREPORT930.Sections("Section4").Controls.Item("LABEL9").Caption = Trim(MACS(17)) & " " & Trim(MACS(18)) & " " & Trim(MACS(19))
        
        Set rs = conn.Execute("Select DES4 from PLEADINGS where Pcode=" & MLET)
        str = Trim(rs.Fields(0))
        
        
        DATAREPORT930.Sections("Section6").Controls.Item("LABEL8").Caption = str
        'DATAREPORT930.Sections("Section6").Controls.Item("LABEL7").Caption = "*" & Trim(MCASE) & "*"
        DATAREPORT930.Sections("Section1").Visible = False
        DATAREPORT930.Sections("Section2").Visible = False
        DATAREPORT930.Sections("Section3").Visible = False
        'DATAREPORT930.Sections("Section6").Visible = False
        DATAREPORT930.Sections("Section5").Visible = False
        DATAREPORT930.Sections("Section7").Visible = False
        DATAREPORT930.Sections("Section8").Visible = False
        DATAREPORT930.Sections("Section9").Visible = False
        DATAREPORT930.Sections("Section11").Visible = False
        'DATAREPORT934.Sections("Section10").Visible = False
        
        
        'DATAREPORT930.Show (1)
        
        'MsgBox str
        'DATAREPORT934.Sections("Section4").Controls.Item("LABEL9").Visible = False
        'DATAREPORT934.Sections("Section4").Controls.Item("LABEL1").Top = 0
        'DATAREPORT934.Sections("Section4").Controls.Item("LABEL5").Visible = False
        'DATAREPORT934.Sections("Section4").Controls.Item("LABEL6").Visible = False
        'DATAREPORT934.Sections("Section1").Controls.Item("LABEL1").Caption = str
        save_Pleadings DATAREPORT930, Me.Caption
    Case 12:
                
        Dim PDF As New clsPDF
        Dim par(10) As String
        'Dim I As String
        re(2) = "Sunil"
        If Trim(Form12.Text1(7).Text) <> "" Then
            par(1) = "We have not received the answers to the Information Subpoena that your client received on "
            par(1) = par(1) & DT3 & ". Kindly have your client complete the Questionnaire and return it to this office within (5) days."
        Else
            par(1) = "We have not received the answers to the Information Subpoena which "
            par(1) = par(1) & " was delivered to you on " & DT3 & "."
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
        PDF.CreatePDF "C:\temp\LET10.pdf"
        PDF.SetLetterhead "C:\ACSSQL\DOC\LH_Content.txt"
        'PDF.SetLetterheadFooter LetterheadFooterText()            ' bottom disclosure (variable text)
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
        PDF.SetStartX 1.25
        PDF.PrintText re(2)
        PDF.PrintText ""
        PDF.PrintText ""
        PDF.PrintText Trim(Text1(1).Text), pdfLeft
        PDF.PrintText ""
        For I = 1 To 10
            If Len(Trim$(par(I))) > 0 Then
                PDF.SetStartX 0.5
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
        
        PDF.SavePDF
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


