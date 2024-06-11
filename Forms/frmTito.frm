VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmViewEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Logs"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4665
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   4755
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&ADD"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&DELETE"
         Height          =   495
         Left            =   1260
         TabIndex        =   14
         Top             =   4050
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "CL&OSE"
         Height          =   495
         Left            =   3510
         TabIndex        =   13
         Top             =   4020
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTito 
         Height          =   3765
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   4485
         _cx             =   7911
         _cy             =   6641
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4194304
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmTito.frx":6852
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSDataListLib.DataCombo cboEmployee 
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   6090
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   315
      Left            =   2130
      TabIndex        =   1
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   90505217
      CurrentDate     =   39034
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   315
      Left            =   4710
      TabIndex        =   2
      Top             =   5250
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   90505217
      CurrentDate     =   39034
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&SHOW"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6510
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo cboPeriod 
      Height          =   315
      Left            =   2130
      TabIndex        =   9
      Top             =   4830
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboBranch 
      Height          =   315
      Left            =   2130
      TabIndex        =   11
      Top             =   5670
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "Branch"
      Height          =   315
      Left            =   450
      TabIndex        =   12
      Top             =   5670
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Payroll period"
      Height          =   195
      Left            =   390
      TabIndex        =   10
      Top             =   4890
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Tito Dates From"
      Height          =   315
      Left            =   390
      TabIndex        =   5
      Top             =   5250
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   315
      Left            =   4110
      TabIndex        =   4
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Specific Employee"
      Height          =   255
      Left            =   390
      TabIndex        =   3
      Top             =   6090
      Width           =   1635
   End
End
Attribute VB_Name = "frmViewEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
