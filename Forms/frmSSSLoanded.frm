VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSSSLoanded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Loan Deduction"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmSSSLoanded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFromMO 
      Height          =   315
      ItemData        =   "frmSSSLoanded.frx":6852
      Left            =   2160
      List            =   "frmSSSLoanded.frx":687A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   2715
   End
   Begin VB.TextBox txtFY 
      Height          =   315
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1710
      TabIndex        =   1
      Top             =   1500
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1500
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1110
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Payroll year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4980
      Y1              =   1020
      Y2              =   1020
   End
End
Attribute VB_Name = "frmSSSLoanded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
