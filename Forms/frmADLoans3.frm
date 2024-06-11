VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADLoans3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Payment"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      TabIndex        =   8
      Top             =   3285
      Width           =   3570
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmADLoans3.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   1830
         TabIndex        =   3
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   688
         Caption         =   "&OK"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmADLoans3.frx":0CDA
         cBack           =   14737632
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3345
      Left            =   0
      TabIndex        =   4
      Top             =   -75
      Width           =   3600
      Begin TDBNumber6Ctl.TDBNumber txtAmntPaid 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   570
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   556
         Calculator      =   "frmADLoans3.frx":19B4
         Caption         =   "frmADLoans3.frx":19D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans3.frx":1A40
         Keys            =   "frmADLoans3.frx":1A5E
         Spin            =   "frmADLoans3.frx":1AA8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,###,##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBNumber6Ctl.TDBNumber txtPreviousBalance 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   225
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   556
         Calculator      =   "frmADLoans3.frx":1AD0
         Caption         =   "frmADLoans3.frx":1AF0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans3.frx":1B5C
         Keys            =   "frmADLoans3.frx":1B7A
         Spin            =   "frmADLoans3.frx":1BC4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   14737632
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,###,##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBNumber6Ctl.TDBNumber txtCurrentBalance 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   915
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   556
         Calculator      =   "frmADLoans3.frx":1BEC
         Caption         =   "frmADLoans3.frx":1C0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans3.frx":1C78
         Keys            =   "frmADLoans3.frx":1C96
         Spin            =   "frmADLoans3.frx":1CE0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   14737632
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,###,##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBText6Ctl.TDBText txtRemarks 
         Height          =   1500
         Left            =   45
         TabIndex        =   10
         Tag             =   "txtRegistrationRemarks"
         Top             =   1785
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   2646
         Caption         =   "frmADLoans3.frx":1D08
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans3.frx":1D74
         Key             =   "frmADLoans3.frx":1D92
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   4210752
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   -1
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   100
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   1530
         Width           =   3420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Current balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   -405
         TabIndex        =   7
         Top             =   960
         Width           =   2025
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount paid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   -405
         TabIndex        =   5
         Top             =   270
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmADLoans3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsDateTime          As ADODB.Recordset
    
    Dim mLoanCode          As Integer
    Dim mLoanDedCode        As Integer

    If Not IsNumeric(txtAmntPaid.Text) Then
        MsgBox "Please enter a number.", vbExclamation + vbOKOnly
        txtAmntPaid.SetFocus
        txtAmntPaid.SelStart = 0
        txtAmntPaid.SelLength = Len(txtAmntPaid.Text)
        Exit Sub
    End If

    If CDbl(txtAmntPaid.Text) = 0 Then
        MsgBox "Amount paid should not be zero.", vbExclamation + vbOKOnly
        txtAmntPaid.SetFocus
        txtAmntPaid.SelStart = 0
        txtAmntPaid.SelLength = Len(txtAmntPaid.Text)
        Exit Sub
    End If

    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    
    With frmADLoans
        
        mLoanDedCode = LastLoanCodeUsed(.rsLoans!loancode)
        
        ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled, " & _
                        "remarks,usercode,payyear,paymonth) values " & _
                      "(" & mLoanDedCode & "," & .rsLoans!loancode & "," & .rsLoans!loantypescode & "," & .mEmployeeCode & "," & _
                       Format(txtAmntPaid.Text, "##0.00") & ",'" & Format(rsDateTime!currentdate, "YYYY-MM-DD") & "', " & Format(CDbl(.rsLoans!loanamnt) - CDbl(txtCurrentBalance.Text), "##0.00") & " ," & Format(txtCurrentBalance.Text, "##0.00") & ",'Y','N'," & _
                       "'" & Swap(txtRemarks.Text) & "'," & GlobalUserID & "," & Format(rsDateTime!currentdate, "YYYY") & ",'" & Format(rsDateTime!currentdate, "MMMM") & "')"

        If CDbl(txtCurrentBalance.Text) = 0 Then
            ConMain.Execute "update loans set status = 'Paid' where loancode = " & CInt(.rsLoans!loancode) & ""
        ElseIf CDbl(txtCurrentBalance.Text) < 0 Then
            ConMain.Execute "update loans set status = 'Over Paid' where loancode = " & CInt(.rsLoans!loancode) & ""
        End If


    End With
    ConMain.CommitTrans
    
    Me.MousePointer = vbHourglass
    With frmADLoans
        mLoanCode = .rsLoans!loancode
        .rsLoans.Requery
        .Get_LoanSum
        .rsLoans.MoveFirst
        .rsLoans.Find "loancode = " & mLoanCode & ""
    End With
    Me.MousePointer = vbDefault
    Unload Me
    

End Sub

Private Sub Form_Activate()

    txtAmntPaid.SetFocus

End Sub

Private Sub Form_Load()

    Dim rsLoanDed           As ADODB.Recordset

    With frmADLoans

        NetOpen rsLoanDed, "select balance from loanded where loancode = " & .rsLoans!loancode & " and fnlz = 'Y' and cancelled = 'N' order by loandedcode desc limit 1"

        txtPreviousBalance.Text = Format(rsLoanDed!balance, "#,##0.00")
        txtCurrentBalance.Text = Format(rsLoanDed!balance, "#,##0.00")

    End With


End Sub

Private Sub txtAmntPaid_GotFocus()
    With txtAmntPaid
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAmntPaid_LostFocus()
    If IsNumeric(txtAmntPaid.Text) Then
        If CDbl(txtAmntPaid.Text) > CDbl(txtPreviousBalance.Text) Then
            txtAmntPaid.Text = txtPreviousBalance.Text
        End If
        txtCurrentBalance.Text = Format(CDbl(txtPreviousBalance.Text) - CDbl(txtAmntPaid.Text), "#,##0.00")
    End If
End Sub

Private Sub txtAmntPaid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
End Sub
