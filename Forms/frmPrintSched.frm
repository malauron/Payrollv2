VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintSched 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Schedule"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmPrintSched.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Criteria"
      Height          =   2955
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6555
      Begin MSDataListLib.DataCombo cboBranch 
         Height          =   315
         Left            =   3180
         TabIndex        =   3
         Top             =   360
         Width           =   3195
         _ExtentX        =   5636
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
      Begin MSDataListLib.DataCombo cboDivision 
         Height          =   315
         Left            =   3180
         TabIndex        =   4
         Top             =   780
         Width           =   3195
         _ExtentX        =   5636
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
      Begin MSDataListLib.DataCombo cboEmployee 
         Height          =   315
         Left            =   3180
         TabIndex        =   5
         Top             =   2460
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo cboPeriod 
         Height          =   315
         Left            =   3180
         TabIndex        =   6
         Top             =   1200
         Width           =   3195
         _ExtentX        =   5636
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
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Top             =   1620
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90505217
         CurrentDate     =   39038
      End
      Begin MSComCtl2.DTPicker dtpTodate 
         Height          =   315
         Left            =   3180
         TabIndex        =   8
         Top             =   2040
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90505217
         CurrentDate     =   39038
      End
      Begin VB.Label Label2 
         Caption         =   "Work Dates"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Branches"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Payroll period"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Specific Employee"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Divisions"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "From "
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Until"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   2100
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Report"
      Height          =   555
      Left            =   3600
      TabIndex        =   1
      Top             =   3420
      Width           =   1395
   End
   Begin VB.CommandButton cmdCLose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   555
      Left            =   5160
      TabIndex        =   0
      Top             =   3420
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   3060
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPrintSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
