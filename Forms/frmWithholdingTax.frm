VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmWithholdingTax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withholding Tax Table"
   ClientHeight    =   6765
   ClientLeft      =   3915
   ClientTop       =   2325
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWithholdingTax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tlbWhtax 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
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
   Begin VB.Frame Frame1 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   8175
      Begin VB.CommandButton cmdPrint 
         Height          =   285
         Left            =   7620
         Picture         =   "frmWithholdingTax.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdGrid 
         Height          =   285
         Left            =   7080
         Picture         =   "frmWithholdingTax.frx":6BDC
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   285
         Left            =   6000
         Picture         =   "frmWithholdingTax.frx":D42E
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   285
         Left            =   6540
         Picture         =   "frmWithholdingTax.frx":13C80
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   300
         Width           =   495
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt01 
         Height          =   375
         Left            =   1200
         TabIndex        =   39
         Top             =   1320
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin MSDataListLib.DataCombo cbobrcktdesc 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   780
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtbrcktdesc 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   4395
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt02 
         Height          =   375
         Left            =   1200
         TabIndex        =   40
         Top             =   1800
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt03 
         Height          =   375
         Left            =   1200
         TabIndex        =   41
         Top             =   2280
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt04 
         Height          =   375
         Left            =   1200
         TabIndex        =   42
         Top             =   2760
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt05 
         Height          =   375
         Left            =   1200
         TabIndex        =   43
         Top             =   3180
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt06 
         Height          =   375
         Left            =   1200
         TabIndex        =   44
         Top             =   3660
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt07 
         Height          =   375
         Left            =   1200
         TabIndex        =   45
         Top             =   4140
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt08 
         Height          =   375
         Left            =   1200
         TabIndex        =   46
         Top             =   4620
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt09 
         Height          =   375
         Left            =   1200
         TabIndex        =   47
         Top             =   5100
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtbrckt10 
         Height          =   375
         Left            =   1200
         TabIndex        =   48
         Top             =   5580
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor1 
         Height          =   375
         Left            =   3720
         TabIndex        =   49
         Top             =   1320
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor2 
         Height          =   375
         Left            =   3720
         TabIndex        =   50
         Top             =   1800
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor3 
         Height          =   375
         Left            =   3720
         TabIndex        =   51
         Top             =   2280
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor4 
         Height          =   375
         Left            =   3720
         TabIndex        =   52
         Top             =   2760
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor5 
         Height          =   375
         Left            =   3720
         TabIndex        =   53
         Top             =   3180
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor6 
         Height          =   375
         Left            =   3720
         TabIndex        =   54
         Top             =   3660
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor7 
         Height          =   375
         Left            =   3720
         TabIndex        =   55
         Top             =   4140
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor8 
         Height          =   375
         Left            =   3720
         TabIndex        =   56
         Top             =   4620
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor9 
         Height          =   375
         Left            =   3720
         TabIndex        =   57
         Top             =   5100
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtfactor10 
         Height          =   375
         Left            =   3720
         TabIndex        =   58
         Top             =   5580
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon1 
         Height          =   375
         Left            =   6360
         TabIndex        =   59
         Top             =   1320
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon2 
         Height          =   375
         Left            =   6360
         TabIndex        =   60
         Top             =   1800
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon3 
         Height          =   375
         Left            =   6360
         TabIndex        =   61
         Top             =   2280
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon4 
         Height          =   375
         Left            =   6360
         TabIndex        =   62
         Top             =   2760
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon5 
         Height          =   375
         Left            =   6360
         TabIndex        =   63
         Top             =   3180
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon6 
         Height          =   375
         Left            =   6360
         TabIndex        =   64
         Top             =   3660
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon7 
         Height          =   375
         Left            =   6360
         TabIndex        =   65
         Top             =   4140
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon8 
         Height          =   375
         Left            =   6360
         TabIndex        =   66
         Top             =   4620
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon9 
         Height          =   375
         Left            =   6360
         TabIndex        =   67
         Top             =   5100
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtaddon10 
         Height          =   375
         Left            =   6360
         TabIndex        =   68
         Top             =   5580
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtexemption 
         Height          =   375
         Left            =   4440
         TabIndex        =   69
         Top             =   300
         Width           =   1515
         _extentx        =   2672
         _extenty        =   661
      End
      Begin MSDataListLib.DataCombo cbobrcktno 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   300
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtbrcktno 
         Height          =   315
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   3
         Top             =   300
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 10"
         Height          =   315
         Index           =   29
         Left            =   5400
         TabIndex        =   37
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 9"
         Height          =   315
         Index           =   28
         Left            =   5400
         TabIndex        =   36
         Top             =   5220
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 8"
         Height          =   315
         Index           =   27
         Left            =   5400
         TabIndex        =   35
         Top             =   4740
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 7"
         Height          =   315
         Index           =   26
         Left            =   5400
         TabIndex        =   34
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 6"
         Height          =   315
         Index           =   25
         Left            =   5400
         TabIndex        =   33
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 5"
         Height          =   315
         Index           =   24
         Left            =   5400
         TabIndex        =   32
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 4"
         Height          =   315
         Index           =   23
         Left            =   5400
         TabIndex        =   31
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 3"
         Height          =   315
         Index           =   22
         Left            =   5400
         TabIndex        =   30
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 2"
         Height          =   315
         Index           =   21
         Left            =   5400
         TabIndex        =   29
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Add-On 1"
         Height          =   315
         Index           =   20
         Left            =   5400
         TabIndex        =   28
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 10"
         Height          =   315
         Index           =   19
         Left            =   2820
         TabIndex        =   27
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 9"
         Height          =   315
         Index           =   18
         Left            =   2820
         TabIndex        =   26
         Top             =   5220
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 8"
         Height          =   315
         Index           =   17
         Left            =   2820
         TabIndex        =   25
         Top             =   4740
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 7"
         Height          =   315
         Index           =   16
         Left            =   2820
         TabIndex        =   24
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 6"
         Height          =   315
         Index           =   15
         Left            =   2820
         TabIndex        =   23
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 5"
         Height          =   315
         Index           =   14
         Left            =   2820
         TabIndex        =   22
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 4"
         Height          =   315
         Index           =   13
         Left            =   2820
         TabIndex        =   21
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 3"
         Height          =   315
         Index           =   12
         Left            =   2820
         TabIndex        =   20
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 2"
         Height          =   315
         Index           =   11
         Left            =   2820
         TabIndex        =   19
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Factor 1"
         Height          =   315
         Index           =   10
         Left            =   2820
         TabIndex        =   18
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 10"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 9"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   5220
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 8"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   4740
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 7"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 6"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 5"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 4"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 3"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 2"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bracket 1"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Exemption"
         Height          =   255
         Left            =   3300
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   795
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   2580
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
            Picture         =   "frmWithholdingTax.frx":1A4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWithholdingTax.frx":20D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWithholdingTax.frx":27596
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWithholdingTax.frx":2DDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWithholdingTax.frx":3465A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWithholdingTax.frx":3AEBC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWithholdingTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
