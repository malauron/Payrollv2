VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSSSTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Table"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSSSTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   6255
      Begin VB.CommandButton cmdPrint 
         Height          =   285
         Left            =   5460
         Picture         =   "frmSSSTable.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdGrid 
         Height          =   285
         Left            =   4860
         Picture         =   "frmSSSTable.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   285
         Left            =   3660
         Picture         =   "frmSSSTable.frx":D42E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   285
         Left            =   4260
         Picture         =   "frmSSSTable.frx":13C80
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtbrckt01 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox txtbrckt02 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtsalcrdt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox txtssser 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtsssee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   2580
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo cbossscode 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtssscode 
         Height          =   315
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "E. C."
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   3060
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "EmployEE's Share"
         Height          =   315
         Index           =   9
         Left            =   600
         TabIndex        =   14
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "EmployER's Share"
         Height          =   315
         Index           =   8
         Left            =   600
         TabIndex        =   13
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Salary Credit"
         Height          =   315
         Index           =   7
         Left            =   600
         TabIndex        =   12
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 2"
         Height          =   315
         Index           =   6
         Left            =   600
         TabIndex        =   11
         Top             =   1380
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 1"
         Height          =   315
         Index           =   5
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Code"
         Height          =   315
         Index           =   4
         Left            =   600
         TabIndex        =   9
         Top             =   540
         Width           =   1275
      End
   End
   Begin MSComctlLib.Toolbar tlbSSS 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":1A4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":20D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":27596
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":2DDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":3465A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSSSTable.frx":3AEBC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSSSTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
