VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSchedule2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit manpower schedule"
   ClientHeight    =   4800
   ClientLeft      =   3495
   ClientTop       =   4800
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedule2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Overtime after regular hours"
      Height          =   855
      Left            =   60
      TabIndex        =   16
      Top             =   1560
      Width           =   7995
      Begin MSComCtl2.DTPicker mSt3in 
         Height          =   315
         Left            =   1830
         TabIndex        =   17
         Top             =   390
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90505218
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker mSt3out 
         Height          =   315
         Left            =   5730
         TabIndex        =   18
         Top             =   390
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90505218
         CurrentDate     =   39039
      End
      Begin VB.Label Label5 
         Caption         =   "Overtime In"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Overtime Out"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkBrokenOt 
      Caption         =   "Broken OT"
      Height          =   315
      Left            =   4980
      TabIndex        =   15
      Top             =   3660
      Width           =   3135
   End
   Begin CitronSoftwarePayroll.ucTextBox txtgrantothrs 
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   2580
      Width           =   2175
      _extentx        =   3836
      _extenty        =   661
   End
   Begin VB.CheckBox Check2 
      Caption         =   "On official travel"
      Height          =   255
      Left            =   2220
      TabIndex        =   8
      Top             =   3660
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   4260
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   4260
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   4260
      Width           =   1335
   End
   Begin VB.TextBox txtReason 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   3180
      Width           =   5055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dayoff"
      Height          =   255
      Left            =   308
      TabIndex        =   7
      Top             =   3660
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo dbcShiftLook 
      Bindings        =   "frmSchedule2.frx":6852
      Height          =   330
      Left            =   1860
      TabIndex        =   5
      Top             =   1020
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "shiftdesc"
      Text            =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1898
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1020
      Width           =   5775
   End
   Begin VB.Label Label5 
      Caption         =   "Granted OT Hours"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2700
      Width           =   1755
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8160
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Label lblHoltype 
      Caption         =   "lblHoltype"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1875
      TabIndex        =   12
      Top             =   540
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "Reason"
      Height          =   255
      Left            =   105
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Shift schedule"
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Holiday"
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label lblWorkdate 
      Caption         =   "lblWorkdate"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1875
      TabIndex        =   1
      Top             =   180
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Work date"
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmSchedule2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
