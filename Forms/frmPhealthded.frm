VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhealthded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Philhealth contribution summary"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhealthded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboToMO 
      Height          =   315
      ItemData        =   "frmPhealthded.frx":6852
      Left            =   2220
      List            =   "frmPhealthded.frx":687A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1110
      Width           =   2715
   End
   Begin VB.ComboBox cboFromMO 
      Height          =   315
      ItemData        =   "frmPhealthded.frx":68E0
      Left            =   2220
      List            =   "frmPhealthded.frx":6908
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   690
      Width           =   2715
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
      Left            =   3420
      TabIndex        =   2
      Top             =   2070
      Width           =   1575
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
      Left            =   1740
      TabIndex        =   1
      Top             =   2070
      Width           =   1575
   End
   Begin VB.TextBox txtFY 
      Height          =   315
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   1680
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   1590
      Y2              =   1590
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
      Left            =   390
      TabIndex        =   4
      Top             =   735
      Width           =   1455
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
      Left            =   390
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmPhealthded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
