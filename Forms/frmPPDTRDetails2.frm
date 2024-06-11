VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPDTRDetails2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin LinkProPayroll.b8ChildTitleBar titlebar 
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   609
      Caption         =   "Title"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   255
      Width           =   5730
      Begin lvButton.lvButtons_H cmdOK 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1965
         TabIndex        =   6
         Top             =   1890
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
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
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin TDBDate6Ctl.TDBDate txtFromDate 
         Height          =   300
         Left            =   1245
         TabIndex        =   0
         Top             =   345
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   529
         Calendar        =   "frmPPDTRDetails2.frx":0000
         Caption         =   "frmPPDTRDetails2.frx":0106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPPDTRDetails2.frx":016C
         Keys            =   "frmPPDTRDetails2.frx":018A
         Spin            =   "frmPPDTRDetails2.frx":01E8
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "09/29/2008"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39720
         CenturyMode     =   0
      End
      Begin TDBTime6Ctl.TDBTime txt1stIn 
         Height          =   300
         Left            =   1455
         TabIndex        =   2
         Top             =   855
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   529
         Caption         =   "frmPPDTRDetails2.frx":0210
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmPPDTRDetails2.frx":0276
         Spin            =   "frmPPDTRDetails2.frx":02C6
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "10:25"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.434027777777778
      End
      Begin TDBDate6Ctl.TDBDate txtToDate 
         Height          =   300
         Left            =   3180
         TabIndex        =   1
         Top             =   345
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   529
         Calendar        =   "frmPPDTRDetails2.frx":02EE
         Caption         =   "frmPPDTRDetails2.frx":03F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPPDTRDetails2.frx":045A
         Keys            =   "frmPPDTRDetails2.frx":0478
         Spin            =   "frmPPDTRDetails2.frx":04D6
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "09/29/2008"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39720
         CenturyMode     =   0
      End
      Begin TDBTime6Ctl.TDBTime txt1stOut 
         Height          =   300
         Left            =   3480
         TabIndex        =   3
         Top             =   855
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   529
         Caption         =   "frmPPDTRDetails2.frx":04FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmPPDTRDetails2.frx":0564
         Spin            =   "frmPPDTRDetails2.frx":05B4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "10:25"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.434027777777778
      End
      Begin TDBTime6Ctl.TDBTime txt2ndIn 
         Height          =   300
         Left            =   1455
         TabIndex        =   4
         Top             =   1200
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   529
         Caption         =   "frmPPDTRDetails2.frx":05DC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmPPDTRDetails2.frx":0642
         Spin            =   "frmPPDTRDetails2.frx":0692
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "10:25"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.434027777777778
      End
      Begin TDBTime6Ctl.TDBTime txt2ndOut 
         Height          =   300
         Left            =   3480
         TabIndex        =   5
         Top             =   1200
         Width           =   1020
         _Version        =   65536
         _ExtentX        =   1799
         _ExtentY        =   529
         Caption         =   "frmPPDTRDetails2.frx":06BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmPPDTRDetails2.frx":0720
         Spin            =   "frmPPDTRDetails2.frx":0770
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "10:25"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.434027777777778
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "2nd OUT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   1980
         TabIndex        =   14
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "2nd IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   -75
         TabIndex        =   13
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "1st OUT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   1980
         TabIndex        =   12
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "1st IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   -75
         TabIndex        =   11
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   -270
         TabIndex        =   9
         Top             =   375
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0030A0B8&
         X1              =   285
         X2              =   5460
         Y1              =   1725
         Y2              =   1725
      End
   End
End
Attribute VB_Name = "frmPPDTRDetails2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    
    On Error GoTo ErrHndlr

    If Not IsDate(txtFromDate.Text) Then
        MsgBox "Incorrect date format.", vbExclamation + vbOKOnly
        txtFromDate.SetFocus
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate.Text)
        Exit Sub
    End If
    
    If CDate(Format(txtFromDate.Text, "MM/DD/YYYY")) < CDate(Format(frmPPDTRdetails.tdbPayrollPeriod.Columns("wrkdatefrom").Text, "MM/DD/YYYY")) Or _
        CDate(Format(txtFromDate.Text, "MM/DD/YYYY")) > CDate(Format(frmPPDTRdetails.tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")) Then
        MsgBox "Date is out of range from payroll period.", vbExclamation + vbOKOnly
        txtFromDate.SetFocus
        txtFromDate.SelStart = 0
        txtFromDate.SelLength = Len(txtFromDate.Text)
        Exit Sub
    End If
    
    If Not IsDate(txtToDate.Text) Then
        MsgBox "Incorrect date format.", vbExclamation + vbOKOnly
        txtToDate.SetFocus
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate.Text)
        Exit Sub
    End If
    
    If CDate(Format(txtToDate.Text, "MM/DD/YYYY")) < CDate(Format(frmPPDTRdetails.tdbPayrollPeriod.Columns("wrkdatefrom").Text, "MM/DD/YYYY")) Or _
        CDate(Format(txtToDate.Text, "MM/DD/YYYY")) > CDate(Format(frmPPDTRdetails.tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")) Then
        MsgBox "Date is out of range from payroll period.", vbExclamation + vbOKOnly
        txtToDate.SetFocus
        txtToDate.SelStart = 0
        txtToDate.SelLength = Len(txtToDate.Text)
        Exit Sub
    End If
    
    If Not IsDate(txt1stIn.Text) Then
        If IsDate(txt1stOut.Text) Then
            MsgBox "Incorrect time format.", vbExclamation + vbOKOnly
            txt1stIn.SetFocus
            txt1stIn.SelStart = 0
            txt1stIn.SelLength = Len(txt1stIn.Text)
            Exit Sub
        End If
    End If
    
    If Not IsDate(txt1stOut.Text) Then
        If IsDate(txt1stIn.Text) Then
            MsgBox "Inorrect time format.", vbExclamation + vbOKOnly
            txt1stOut.SetFocus
            txt1stOut.SelStart = 0
            txt1stOut.SelLength = Len(txt1stOut.Text)
            Exit Sub
        End If
    End If
    
    If Not IsDate(txt2ndIn.Text) Then
        If IsDate(txt2ndOut.Text) Then
            MsgBox "Incorrect time format.", vbExclamation + vbOKOnly
            txt2ndIn.SetFocus
            txt2ndIn.SelStart = 0
            txt2ndIn.SelLength = Len(txt2ndIn.Text)
            Exit Sub
        End If
    End If
    
    If Not IsDate(txt2ndOut.Text) Then
        If IsDate(txt2ndIn.Text) Then
            MsgBox "Inorrect time format.", vbExclamation + vbOKOnly
            txt2ndOut.SetFocus
            txt2ndOut.SelStart = 0
            txt2ndOut.SelLength = Len(txt2ndOut.Text)
            Exit Sub
        End If
    End If
    
    With frmPPDTRdetails
            If Not .rsDtrTmp.EOF Then
                .mRow = .tdgDtr.Row
                .mCol = .tdgDtr.Col
                .rsDtrTmp.MoveFirst
                Do While Not .rsDtrTmp.EOF
                    If CDate(.tdgDtr.Columns("wrkdate").Text) >= CDate(Format(txtFromDate.Text, "MM/DD/YYYY")) And CDate(.tdgDtr.Columns("wrkdate").Text) <= CDate(Format(txtToDate.Text, "MM/DD/YYYY")) Then
                        .tdgDtr.Columns("st1in").Text = IIf(IsDate(txt1stIn.Text), Format(txt1stIn.Text, "HH:NN"), "")
                        .tdgDtr.Columns("st1out").Text = IIf(IsDate(txt1stOut.Text), Format(txt1stOut.Text, "HH:NN"), "")
                        .tdgDtr.Columns("st2in").Text = IIf(IsDate(txt2ndIn.Text), Format(txt2ndIn.Text, "HH:NN"), "")
                        .tdgDtr.Columns("st2out").Text = IIf(IsDate(txt2ndOut.Text), Format(txt2ndOut.Text, "HH:NN"), "")
                    End If
                    .rsDtrTmp.MoveNext
                Loop
                .tdgDtr.Row = .mRow
                .tdgDtr.Col = .mCol
            End If
    End With
    
    Unload Me
    
ErrHndlr:
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    txtFromDate.SetFocus
End Sub

Private Sub txtFromDate_GotFocus()
    txtFromDate.SelStart = 0
    txtFromDate.SelLength = Len(txtFromDate.Text)
End Sub

Private Sub txtFromDate_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtToDate_GotFocus()
    txtToDate.SelStart = 0
    txtToDate.SelLength = Len(txtToDate.Text)
End Sub

Private Sub txtToDate_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt1stIn_GotFocus()
    txt1stIn.SelStart = 0
    txt1stIn.SelLength = Len(txt1stIn.Text)
End Sub

Private Sub txt1stIn_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt1stOut_GotFocus()
    txt1stOut.SelStart = 0
    txt1stOut.SelLength = Len(txt1stOut.Text)
End Sub
Private Sub txt1stOut_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt2ndIn_GotFocus()
    txt2ndIn.SelStart = 0
    txt2ndIn.SelLength = Len(txt2ndIn.Text)
End Sub

Private Sub txt2ndIn_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt2ndOut_GotFocus()
    txt2ndOut.SelStart = 0
    txt2ndOut.SelLength = Len(txt2ndOut.Text)
End Sub

Private Sub txt2ndOut_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
