VERSION 5.00
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask8.ocx"
Begin VB.Form frmPPAddTito 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maintain Time-In Time-Out (TITO)"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPPAddTito.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin TDBMask6Ctl.TDBMask mskWorkdate 
      Height          =   300
      Left            =   1290
      TabIndex        =   7
      Top             =   570
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   529
      Caption         =   "frmPPAddTito.frx":6852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmPPAddTito.frx":68B8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   2
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "&&&&&&&&&&"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "TDBMask1__"
      Value           =   "TDBMask1"
   End
   Begin VB.OptionButton optTime 
      BackColor       =   &H80000016&
      Caption         =   "Time Out"
      Height          =   255
      Index           =   1
      Left            =   2550
      TabIndex        =   4
      Top             =   1305
      Width           =   1095
   End
   Begin VB.OptionButton optTime 
      BackColor       =   &H80000016&
      Caption         =   "Time In"
      Height          =   255
      Index           =   0
      Left            =   1305
      TabIndex        =   3
      Top             =   1305
      Value           =   -1  'True
      Width           =   1095
   End
   Begin OsenXPCntrl.OsenXPButton cmdOk 
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      BTYPE           =   8
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483626
      BCOLO           =   -2147483626
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPPAddTito.frx":68FA
      PICN            =   "frmPPAddTito.frx":6916
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Height          =   465
      Left            =   1140
      TabIndex        =   6
      Top             =   0
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      BTYPE           =   8
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483626
      BCOLO           =   -2147483626
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPPAddTito.frx":6EB2
      PICN            =   "frmPPAddTito.frx":6ECE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBMask6Ctl.TDBMask mskTime 
      Height          =   300
      Left            =   1290
      TabIndex        =   8
      Top             =   900
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   529
      Caption         =   "frmPPAddTito.frx":746A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmPPAddTito.frx":74D0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   2
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "&&&&&&&&&&"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "TDBMask1__"
      Value           =   "TDBMask1"
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Caption         =   "Type"
      Height          =   300
      Left            =   45
      TabIndex        =   2
      Top             =   1305
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Caption         =   "Time"
      Height          =   225
      Left            =   45
      TabIndex        =   1
      Top             =   975
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Caption         =   "Work Date"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   1095
   End
End
Attribute VB_Name = "frmPPAddTito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
