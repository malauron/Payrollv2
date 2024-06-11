VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLOBLeave3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Leave Limit"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "modiifie"
      Default         =   -1  'True
      Height          =   375
      Left            =   13155
      TabIndex        =   8
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1200
      Left            =   15
      TabIndex        =   4
      Top             =   -75
      Width           =   6150
      Begin TDBText6Ctl.TDBText txtLeaveType 
         Height          =   300
         Left            =   1275
         TabIndex        =   5
         Top             =   300
         Width           =   4665
         _Version        =   65536
         _ExtentX        =   8229
         _ExtentY        =   529
         Caption         =   "frmLOBLeave3.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLOBLeave3.frx":006C
         Key             =   "frmLOBLeave3.frx":008A
         BackColor       =   14737632
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   -1
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   2
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber txtLeaveLimit 
         Height          =   315
         Left            =   1275
         TabIndex        =   0
         Top             =   630
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmLOBLeave3.frx":00CE
         Caption         =   "frmLOBLeave3.frx":00EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLOBLeave3.frx":0154
         Keys            =   "frmLOBLeave3.frx":0172
         Spin            =   "frmLOBLeave3.frx":01AC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   4210752
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
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
         Height          =   255
         Left            =   -360
         TabIndex        =   7
         Top             =   345
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Limit"
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
         Height          =   255
         Left            =   -375
         TabIndex        =   6
         Top             =   690
         Width           =   1560
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      TabIndex        =   1
      Top             =   1110
      Width           =   6135
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   15
         TabIndex        =   2
         Top             =   60
         Width           =   1995
         _ExtentX        =   3519
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
         Image           =   "frmLOBLeave3.frx":01D4
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   390
         Left            =   2010
         TabIndex        =   3
         Top             =   60
         Width           =   1995
         _ExtentX        =   3519
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
         Image           =   "frmLOBLeave3.frx":0EAE
         cBack           =   14737632
      End
   End
End
Attribute VB_Name = "frmLOBLeave3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mYear      As Integer
Public mEmpno     As Integer
Public mLvCode    As Integer

Public mLvLimit   As Double

Public mLvName    As String

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  
  If Not IsNumeric(txtLeaveLimit.Text) Then
    MsgBox "Invalid number format.", vbExclamation + vbOKOnly
    txtLeaveLimit.SetFocus
    Exit Sub
  End If
  
  ConMain.Execute "set autocommit = 0"
  ConMain.BeginTrans
  ConMain.Execute "delete from lvlimit where payyear = " & mYear & " and employeecode = " & mEmpno & " and leavetypescode = " & mLvCode & ""
  ConMain.Execute "insert into lvlimit(payyear,employeecode,leavetypescode,lvlimit) values (" & _
                  mYear & ", " & mEmpno & "," & mLvCode & "," & txtLeaveLimit.Value & ") "
  ConMain.CommitTrans
  frmLOBLeave.rsLeaveLimit!lvlimit = txtLeaveLimit.Value
  MsgBox "Limit has been successfully updated.", vbInformation + vbOKOnly
  Unload Me
  
End Sub

Private Sub Form_Load()
  txtLeaveType.Text = mLvName
  txtLeaveLimit.Value = mLvLimit
End Sub

Private Sub txtLeaveLimit_GotFocus()
  With txtLeaveLimit
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLeaveLimit_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

