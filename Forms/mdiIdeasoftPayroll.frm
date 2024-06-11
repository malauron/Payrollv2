VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.MDIForm mdiIdeasoftPayroll 
   Appearance      =   0  'Flat
   BackColor       =   &H00F6F8F8&
   Caption         =   "EIGHT2EIGHT Human Resource Management & Payroll System - LinkPro Technologies Inc."
   ClientHeight    =   8085
   ClientLeft      =   3300
   ClientTop       =   1995
   ClientWidth     =   13995
   Icon            =   "mdiIdeasoftPayroll.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13995
      TabIndex        =   2
      Top             =   7440
      Width           =   13995
      Begin lvButton.lvButtons_H cmd 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         Caption         =   "&Button"
         CapAlign        =   2
         BackStyle       =   4
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
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   14737632
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   14737632
      End
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   2460
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   4339
      ButtonWidth     =   2646
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "i32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Payroll Periods"
            Key             =   "mnuPayrollPeriod"
            Object.ToolTipText     =   "Payroll Periods"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sepLoans"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Loans"
            Key             =   "mnuLoanApplications"
            Object.ToolTipText     =   "Loan Applications"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Leaves"
            Key             =   "mnuLeaves"
            Object.ToolTipText     =   "Leaves"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sepOverTime"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Overtime"
            Key             =   "mnuOvertime"
            Object.ToolTipText     =   "Overtime"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Other Earnings"
            Key             =   "mnuOtherEarnings"
            Object.ToolTipText     =   "Other Earnings"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Other Deductions"
            Key             =   "mnuOtherDeductions"
            Object.ToolTipText     =   "Other Deductions"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Loan deductions"
            Key             =   "mnuGenLoan"
            Object.ToolTipText     =   "Loan deductions"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DTR Summary"
            Key             =   "mnuDtrSummary"
            Object.ToolTipText     =   "Generate Daily Time Log"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sepGenPay"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Generate Payroll"
            Key             =   "mnuGenPayroll"
            Object.ToolTipText     =   "Generate Payroll"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Finalize Payroll"
            Key             =   "mnuFinalizePayroll"
            Object.ToolTipText     =   "Finalize Payroll"
            ImageIndex      =   24
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7785
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   14
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "mdiIdeasoftPayroll.frx":0CCA
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3422
            MinWidth        =   882
            Key             =   "ctrusername"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "User Type:"
            TextSave        =   "User Type:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3422
            MinWidth        =   882
            Key             =   "ctrUserType"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Server and Database Name:"
            TextSave        =   "Server and Database Name:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "ctrservername"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5/2/2023"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:37 AM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2655
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1066
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   3840
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4831
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":550B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":61E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":6EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":7B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":8873
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":954D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":A227
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":AF01
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":BBDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":C8B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":D58F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":E269
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":EF43
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":FC1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":108F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":115D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":122AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":12F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":13C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":14939
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":15613
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":162ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":16FC7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   3255
      Top             =   4425
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
            Picture         =   "mdiIdeasoftPayroll.frx":17CA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1823B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":187D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":18B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":18F09
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":192A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1963D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1A04F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1AA61
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1ADFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1B195
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1B52F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1B8C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1C2DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1CCED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1D6FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1E111
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1EB23
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1F535
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":1FF47
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":204E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   3240
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":20A7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":22411
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":23DA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":25735
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":270C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":28A59
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":2A3EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":2BD7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":2D70F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":2F0A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":2FD7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3065F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3133B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":32017
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":32CF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":339CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":346AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":34F87
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4440
      Top             =   4425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":35C61
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":36673
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":37085
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3741F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":377B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":37B53
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":37EED
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":388FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":39311
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":39D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3A735
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3B147
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3BB59
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3C56B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3CB07
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFormList 
      Left            =   3855
      Top             =   4425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3D0A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3EA35
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":3F711
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":410A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":42A35
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":443C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":45D59
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":46A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4770D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":483E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":490C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":49D9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4A67B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4B357
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4C033
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4CD0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4D5F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4E2CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4EBAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":4F887
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":5121B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiIdeasoftPayroll.frx":52BAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMasterData 
      Caption         =   "&Master Data"
      Enabled         =   0   'False
      Begin VB.Menu mnuNetPayCap 
         Caption         =   "Net Pay Cap"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeductPriority 
         Caption         =   "Deductions Priority"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employees"
         Enabled         =   0   'False
         Begin VB.Menu mnuEmployeeMasterfile 
            Caption         =   "Employee Masterfile"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFingerprintRegistration 
            Caption         =   "Employee Fingerprint Registration"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnufingerprintverification 
            Caption         =   "Employee Fingerprint Verification "
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuJobTitle 
         Caption         =   "Job Title"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmployStatus 
         Caption         =   "Employment Status"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShiftSched 
         Caption         =   "Shift Schedules"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu province 
         Caption         =   "Province"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Municipal 
         Caption         =   "Municipal"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu barangay 
         Caption         =   "Barangay"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sepBranches 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branches"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDivision 
         Caption         =   "Divisions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCostCenter 
         Caption         =   "Cost Centers"
         Enabled         =   0   'False
      End
      Begin VB.Menu mSection 
         Caption         =   "Section"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepBank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBank 
         Caption         =   "Banks"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtherEarning 
         Caption         =   "Other Earnings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherDeduct 
         Caption         =   "Other Deductions-Outright"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLoanType 
         Caption         =   "Loan/Advance Types"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherBonusType 
         Caption         =   "Other Bonus Types"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLeaveType 
         Caption         =   "Leave Types"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHoliday 
         Caption         =   "Holidays"
         Enabled         =   0   'False
      End
      Begin VB.Menu leavecredits 
         Caption         =   "Leave Credits"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWTaxTable 
         Caption         =   "Withholding Tax Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSSSTable 
         Caption         =   "SSS Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPHICTable 
         Caption         =   "PhilHealth Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHDMFTable 
         Caption         =   "HDMF Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserGroups 
         Caption         =   "User Groups"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHRM 
      Caption         =   "HR Management"
      Enabled         =   0   'False
      Begin VB.Menu mnuEmployeePerformanceEvaluation 
         Caption         =   "Employee Performance Evaluation"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMealAgreementForm 
         Caption         =   "Meal Agreement Form"
      End
   End
   Begin VB.Menu mnuPayrollProcess 
      Caption         =   "&Payroll Procedures"
      Enabled         =   0   'False
      Begin VB.Menu mnuPayrollPeriod 
         Caption         =   "Payroll Period"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepApplications 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoanApplications 
         Caption         =   "Loan Applications"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLeaves 
         Caption         =   "Leaves"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReceivables 
         Caption         =   "Receivables"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepPayrollProcedure 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOvertime 
         Caption         =   "Overtime"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherEarnings 
         Caption         =   "Other Earnings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherDeductions 
         Caption         =   "Other Deductions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGenLoan 
         Caption         =   "Loan Deductions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDtrSummary 
         Caption         =   "DTR Summary"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenPayroll 
         Caption         =   "Generate Payroll"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFinalizePayroll 
         Caption         =   "Finalize Payroll"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Begin VB.Menu mnuPayslip 
         Caption         =   "Payslips"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPayrollReg 
         Caption         =   "Payroll Register"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuNetPayList 
         Caption         =   "Net Pay Report"
         Enabled         =   0   'False
         Begin VB.Menu mnuBankNetpay 
            Caption         =   "For Banks"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCashPay 
            Caption         =   "Cash Payroll"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuNegativenet 
            Caption         =   "Employees with Negative Net"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuNetPay 
            Caption         =   "Net Pay List"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuBasicPay 
         Caption         =   "Basic Pay"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep16 
         Caption         =   "-"
      End
      Begin VB.Menu mActHrsWrk 
         Caption         =   "Actual Time Log Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDailyTimeLog 
         Caption         =   "Daily Time Log Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAttendanceReport 
         Caption         =   "Attendance Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbsenceLate 
         Caption         =   "Absences, Late and Undertime"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRptLeave 
         Caption         =   "Leave"
         Enabled         =   0   'False
         Begin VB.Menu mnuLeaveAvail 
            Caption         =   "Leave Availments"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sptr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOvertimeRep 
         Caption         =   "Overtime "
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOtherEarnRep 
         Caption         =   "Other Earnings"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep50 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatutoryDeduct 
         Caption         =   "Statutory Deductions"
         Enabled         =   0   'False
         Begin VB.Menu mnuSSSContributions 
            Caption         =   "SSS Contributions"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPhilHealthContribution 
            Caption         =   "PhilHealth Contribution"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuHDMFContributions 
            Caption         =   "HDMF Contributions"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuWithholdingTaxContribution 
            Caption         =   "Withholding Tax Contribution"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuOtherDeductRep 
         Caption         =   "Other Deductions"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDeductionList 
         Caption         =   "Deductions List"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLoanRep 
         Caption         =   "Loan Collections/Deductions"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlphalist 
         Caption         =   "Alphalist"
         Enabled         =   0   'False
         Begin VB.Menu mnuSSSEPF 
            Caption         =   "SSS Employee Pre-validation File"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployList 
         Caption         =   "Employee Listing"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBiometric 
         Caption         =   "Biometric ID"
         Enabled         =   0   'False
      End
      Begin VB.Menu SepRptVoucher 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoucherS 
         Caption         =   "Vouchers"
         Enabled         =   0   'False
         Begin VB.Menu mnuVoucherSlips 
            Caption         =   "Voucher Slips"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuVoucherSummary 
            Caption         =   "Voucher Summary"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuAcknowledgmentReceipts 
         Caption         =   "Acknowledgment Receipts"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepEmployeeEvalRep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployeePerformanceEvaluationReport 
         Caption         =   "Employee Performance Evaluation Report"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuMiscellaneous 
      Caption         =   "M&iscellaneous"
      Enabled         =   0   'False
      Begin VB.Menu mnuExportAlphaList 
         Caption         =   "Export Alphalist"
         Enabled         =   0   'False
      End
      Begin VB.Menu SepMnuVoucher 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoucher 
         Caption         =   "Voucher"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCoupon 
         Caption         =   "Coupon"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Enabled         =   0   'False
      Begin VB.Menu mnuLogin 
         Caption         =   "Log-in User"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log-off User"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaintainUser 
         Caption         =   "Maintain Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChangePwd 
         Caption         =   "Change Password"
         Enabled         =   0   'False
      End
      Begin VB.Menu mDLimg 
         Caption         =   "Extract Employees' Picture and Signature"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpdates 
         Caption         =   "Check for updates"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParameter 
         Caption         =   "Parameters"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSignatories 
         Caption         =   "Signatories"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompanyInfo 
         Caption         =   "Company Information"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "mdiIdeasoftPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Const MF_BYPOSITION = &H400&

Dim fn() As String

Private Sub barangay_Click()
  frmMDBarangay.Show
  frmMDBarangay.ZOrder
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Frm As Form
    For Each Frm In Forms
        If LCase(Trim(Frm.Name)) <> LCase(Trim(mdiIdeasoftPayroll.Name)) Then
            If LCase(Trim(Frm.Name)) = LCase(Trim(cmd(Index).Tag)) Then
                Frm.SetFocus
                Frm.ZOrder
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub EmpLogs_Click()
    If allowedAccess("frmUtilEmpLogs") = False Then Exit Sub
    frmUtilEmpLogs.Show
    frmUtilEmpLogs.ZOrder
End Sub

Private Sub LeaveCredits_Click()
    If allowedAccess("frmLeaveCredits") = False Then Exit Sub
    frmLeaveCredits.Show
    frmLeaveCredits.ZOrder
End Sub

Private Sub mActHrsWrk_Click()
    If allowedAccess("frmRptActHrsWrk") = False Then Exit Sub
    frmRptActHrsWrk.Show
    frmRptActHrsWrk.ZOrder
End Sub

Private Sub MDIForm_Load()
        
    mForm_Count = 0
    
    With StatusBar2
        .Panels("ctrusername").Text = Trim(UserName)
        .Panels("ctrusername").AutoSize = sbrContents
        .Panels("ctrUserType").AutoSize = sbrContents
        .Panels("ctrservername").Text = Trim((SQLServerName & " - " & SQLDatabase))
        .Panels("ctrservername").AutoSize = sbrContents
    End With
    
    Me.mnuLogin.Visible = False
    Me.mnuChangePwd.Visible = False
    
    Dim mnuItem As Control
    Dim Btn As Button
'    Dim BM As ButtonMenu
      
    If GlobalUserID = 1 Then
    
      For Each mnuItem In Me.Controls
          If TypeOf mnuItem Is Menu Then
              mnuItem.Enabled = True
          End If
      Next
      
      
      For Each Btn In tlbMenu.Buttons
        Btn.Enabled = True
      Next
    
    Else
      
      Dim rsMenu As New ADODB.Recordset
      
      NetOpen rsMenu, "select * from usergroup_restrictions " & _
                      "where usergroup_id = " & GlobalUserGroupID & " and module = '" & ModuleVersion & "'"
      
      With rsMenu
        If .RecordCount Then
          
          .MoveFirst
          Do While Not .EOF
            
            For Each mnuItem In Me.Controls
                If TypeOf mnuItem Is Menu Then
                  If mnuItem.Caption <> "-" Then
                    If mnuItem.Name = !menuname Then
                      mnuItem.Enabled = True
                      Exit For
                    End If
                  End If
                End If
            Next
            
            .MoveNext
          Loop
          
          .MoveFirst
          Do While Not .EOF
          
            For Each Btn In tlbMenu.Buttons
              If Btn.Key = !menuname Then
                Btn.Enabled = True
                Exit For
              End If
            Next
            
            .MoveNext
          Loop
          
        End If
      End With
      
      mnuMasterData.Enabled = True
      mnuExit.Enabled = True
      
    End If
        
    frmLogo.Show
    
    frmLogo.ZOrder
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Are you sure you want to close this program?", vbYesNo + vbQuestion) = vbYes Then
        CloseOpenForms
        End
    Else
        Cancel = True
    End If

End Sub

Sub CloseOpenForms()

    Dim intFrmNum As Integer
    
    intFrmNum = Forms.count
    
    If intFrmNum <> 1 Then
        Do Until intFrmNum = 1
            Unload Forms(intFrmNum - 1)
            intFrmNum = intFrmNum - 1
        Loop
    End If
    
End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next

    Dim i           As Integer
    Dim mLeft       As Long
    
    If Me.Height < 11190 Or Me.Width < 15480 Then
      Me.Height = 11190
      Me.Width = 15480
    End If
    
    mLeft = 0
    
    For i = 0 To cmd.UBound
        If cmd(i).Visible = True Then
            cmd(i).Top = 0
            cmd(i).Width = (Me.ScaleWidth / (mForm_Count))
            cmd(i).Left = mLeft
            mLeft = mLeft + cmd(i).Width
        End If
    Next
    
End Sub

Private Sub mDLimg_Click()
'    frmUtilDLImg.Show
'    frmUtilDLImg.ZOrder
    
    Dim FSO As FileSystemObject
    Dim rs As ADODB.Recordset
    Dim mystream As ADODB.Stream


    
    If MsgBox("Do you want to extract employees' picture and signature? ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    Set rs = New ADODB.Recordset
    
    Set mystream = New ADODB.Stream

    mystream.Type = adTypeBinary
    
    
    Set FSO = New FileSystemObject
    If Not Dir(App.Path & "\EmpPics", vbDirectory) = vbNullString Then
      FSO.DeleteFolder App.Path & "\EmpPics", True
    End If
     
    Set FSO = New FileSystemObject
    If Not Dir(App.Path & "\EmpSig", vbDirectory) = vbNullString Then
      FSO.DeleteFolder App.Path & "\EmpSig", True
    End If
    
    MkDir App.Path & "\EmpPics"
    MkDir App.Path & "\EmpSig"
    
    NetOpen rs, "select * from emppics"
  
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      Do While Not rs.EOF
        mystream.Open
        mystream.Write rs!images
        mystream.SaveToFile App.Path & "\EmpPics\" & rs!FileName
        mystream.Close
        rs.MoveNext
      Loop
    End If
      
    NetOpen rs, "select * from empsig"
  
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      Do While Not rs.EOF
        mystream.Open
        mystream.Write rs!images
        mystream.SaveToFile App.Path & "\EmpSig\" & rs!FileName
        mystream.Close
        rs.MoveNext
      Loop
    End If
    
    MsgBox "Process completed successfully!", vbInformation + vbOKOnly
    
End Sub

Private Sub mnuAbsenceLate_Click()
    If allowedAccess("frmRptAbsences") = False Then Exit Sub
    frmRptAbsences.Show
    frmRptAbsences.ZOrder
End Sub

Private Sub mnuAcknowledgmentReceipts_Click()
    If allowedAccess("frmRptAcknowledgmentReceipt") = False Then Exit Sub
    frmRptAcknowledgmentReceipt.Show
    frmRptAcknowledgmentReceipt.ZOrder
End Sub

Private Sub mnuBank_Click()
  If allowedAccess("frmMDBanks") = False Then Exit Sub
  frmMDBanks.Show
  frmMDBanks.ZOrder
End Sub

Private Sub mnuBankNetpay_Click()
    If allowedAccess("frmRptNetpayListBank") = False Then Exit Sub
    frmRptNetpayListBank.Show
    frmRptNetpayListBank.ZOrder
End Sub

Private Sub mnuBranch_Click()
    If allowedAccess("frmMDBranch") = False Then Exit Sub
    With frmMDBranch
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuCashPay_Click()
    If allowedAccess("frmRptNetPayListCash") = False Then Exit Sub
    frmRptNetPayListCash.Show
End Sub

Private Sub mnuCheckUpdates_Click()

    Dim FileName1       As String
    Dim txt             As String
    Dim Phrase          As String
    Dim Char1           As String
    Dim mEncrypted      As String
    Dim mSourcePath     As String
    Dim mString         As String
    
    Dim Position        As Integer
    Dim CTR             As Integer
    
    Dim Asc1            As Long
    
    Dim mToUpdate       As Boolean
    
    
    
    If Not DirExists(App.Path & "\Updater.exe") Then
        MsgBox "File not found. (Updater.exe)", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Dir(App.Path & "\ExtData\upd8.dat", vbNormal) <> vbNullString Then
        
        FileName1 = App.Path & "\ExtData\upd8.dat"
    
        mEncrypted = ""
        
        Open FileName1 For Input As #1
            
            Do Until EOF(1)
                Input #1, txt
                mEncrypted = mEncrypted & txt
            Loop
            
        Close #1
        
        Phrase = mEncrypted
        
        mEncrypted = ""
        
        CTR = 0
        
        mToUpdate = False
        
        For Position = Len(Phrase) To 1 Step -1
            Char1 = Mid$(Phrase, Position, 1)
            Asc1 = Asc(Char1)
            Asc1 = (((Asc1 * Asc1) / 2) / 2)
            Asc1 = Sqr(Asc1)
            Char1 = Chr$(Asc1)
            '----
            If Char1 = "," Then
                Char1 = ""
                CTR = CTR + 1
                If CTR = 1 Then
                    mSourcePath = mEncrypted
                    mEncrypted = ""
                ElseIf CTR = 2 Then
                    If mEncrypted = 1 Then
                        mToUpdate = True
                    End If
                    mEncrypted = ""
                End If
            End If
            '-----
            mEncrypted = mEncrypted & Char1
        Next
        
    End If
    
    If Trim(mSourcePath) = "" Then mSourcePath = App.Path & "\" & App.EXEName & ".exe"
    
    Phrase = mSourcePath & "," & 1 & "," & App.EXEName & ".exe" & ","
    mString = ""
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        Asc1 = Asc(Char1)
        Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
        Char1 = Chr$(Asc1)
        mString = mString & Char1
    Next
    
    FileName1 = App.Path & "\Extdata\upd8.dat"
    
    Open FileName1 For Output As #1
        Print #1, mString
    Close #1
    
    Shell (App.Path & "\Updater.exe")
    
    DestroyAllObjects
    
    End
    
End Sub

Private Sub mnuCompanyInfo_Click()
    If allowedAccess("frmUtilCompanyInfo") = False Then Exit Sub
    frmUtilCompanyInfo.Show vbModal
End Sub

Private Sub mnuCostCenter_Click()
    If allowedAccess("frmMDCostCenter") = False Then Exit Sub
    With frmMDCostCenter
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuDailySched_Click()
    If allowedAccess("frmPPSched") = False Then Exit Sub
    frmPPSched.Show
    frmPPSched.ZOrder
End Sub

Private Sub mnuCoupon_Click()
    If allowedAccess("frmAdCoupon") = False Then Exit Sub
    frmAdCoupon.Show
    frmAdCoupon.ZOrder
End Sub

Private Sub mnuDailyTimeLog_Click()
    If allowedAccess("frmRptDailyTimeLog") = False Then Exit Sub
    frmRptDailyTimeLog.Show
    frmRptDailyTimeLog.ZOrder
End Sub

Private Sub mnuDeductPriority_Click()
    If allowedAccess("frmMDDedPriority") = False Then Exit Sub
    frmMDDedPriority.Show
    frmMDDedPriority.ZOrder
End Sub

Private Sub mnuDivision_Click()
    If allowedAccess("frmMDDivision") = False Then Exit Sub
    With frmMDDivision
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuDtrSummary_Click()
    If allowedAccess("frmPPDTR") = False Then Exit Sub
    frmPPDTR.Show
    frmPPDTR.ZOrder
End Sub



Private Sub mnuEmployeeMasterfile_Click()
    If allowedAccess("frmMDEmployee") = False Then Exit Sub
    frmMDEmployee.Show
    frmMDEmployee.ZOrder

End Sub

Private Sub mnuEmployeePerformanceEvaluation_Click()
    If allowedAccess("frmAdEmployeePerformanceEvaluation") = False Then Exit Sub
  frmAdEmployeePerformanceEvaluation.Show
  frmAdEmployeePerformanceEvaluation.ZOrder
End Sub

Private Sub mnuEmployeePerformanceEvaluationReport_Click()
    If allowedAccess("frmRptEmployeePerformanceEvaluation") = False Then Exit Sub
  frmRptEmployeePerformanceEvaluation.Show
  frmRptEmployeePerformanceEvaluation.ZOrder
End Sub

Private Sub mnuEmployStatus_Click()
    If allowedAccess("frmMDStatus") = False Then Exit Sub
    frmMDStatus.Show
    frmMDStatus.ZOrder
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExportAlphaList_Click()
    If allowedAccess("frmRptAlphaList") = False Then Exit Sub
    frmRptAlphaList.Show
    frmRptAlphaList.ZOrder
End Sub

Private Sub mnuFinalizePayroll_Click()
    If allowedAccess("frmPPFinalizePayroll") = False Then Exit Sub
    frmPPFinalizePayroll.Show vbModal
End Sub

Private Sub mnuFingerprintRegistration_Click()
    If allowedAccess("frmMDFingerPrintRegistration") = False Then Exit Sub
  frmMDFingerPrintRegistration.Show vbModal
End Sub

Private Sub mnufingerprintverification_Click()
    If allowedAccess("frmMDFingerPrintVerification") = False Then Exit Sub
  frmMDFingerPrintVerification.Show vbModal
End Sub

Private Sub mnuGenLoan_Click()
    If allowedAccess("frmPPGenLoanDed") = False Then Exit Sub
    frmPPGenLoanDed.Show
    frmPPGenLoanDed.ZOrder
End Sub

Private Sub mnuGenPayroll_Click()
    If allowedAccess("frmPPGeneratePayroll") = False Then Exit Sub
    frmPPGeneratePayroll.Show vbModal
End Sub

Private Sub mnuHDMFContributions_Click()
    If allowedAccess("frmRptHDMFContributions") = False Then Exit Sub
    With frmRptHDMFContributions
        .Show
    End With
End Sub

Private Sub mnuHDMFTable_Click()
    If allowedAccess("frmMDPagibig") = False Then Exit Sub
    With frmMDPagibig
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuHoliday_Click()
    If allowedAccess("frmMDHoliday") = False Then Exit Sub
    With frmMDHoliday
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuImportAyala_Click()
    If allowedAccess("frmUtilImportAll") = False Then Exit Sub
    frmUtilImportAll.Show vbModal
End Sub

Private Sub mnuJobTitle_Click()
    If allowedAccess("frmMDJobTitle") = False Then Exit Sub
    With frmMDJobTitle
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuLeaveCredit_Click()
'    frmRptLeave.Show
'    frmRptLeave.ZOrder
End Sub

Private Sub mnuLeaveAvail_Click()
    If allowedAccess("frmRptLeaveApplications") = False Then Exit Sub
    frmRptLeaveApplications.Show
    frmRptLeaveApplications.ZOrder
End Sub

Private Sub mnuLeaves_Click()
    If allowedAccess("frmLOBLeave") = False Then Exit Sub
    frmLOBLeave.Show
    frmLOBLeave.ZOrder
End Sub

Private Sub mnuLeaveType_Click()
    If allowedAccess("frmMDLeave") = False Then Exit Sub
    With frmMDLeave
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuLoanApplications_Click()
    If allowedAccess("frmADLoans") = False Then Exit Sub
    frmADLoans.Show
    frmADLoans.ZOrder
End Sub

Private Sub mnuLoanRep_Click()
    If allowedAccess("frmRptLoanCollection") = False Then Exit Sub
    frmRptLoanCollection.Show
    frmRptLoanCollection.ZOrder
End Sub

Private Sub mnuLoanType_Click()
    If allowedAccess("frmMDLoanTypes") = False Then Exit Sub
    With frmMDLoanTypes
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnuLogoff_Click()
    If MsgBox("Do you want to log off?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    mLogOn = False
    Unload Me
End Sub

Private Sub mnuMaintainUser_Click()
    If allowedAccess("frmSysUsersInfo") = False Then Exit Sub
    frmSysUsersInfo.Show
End Sub

Private Sub mnuMealAgreementForm_Click()
    If allowedAccess("frmRptMealAgreementForm") = False Then Exit Sub
  frmRptMealAgreementForm.Show
  frmRptMealAgreementForm.ZOrder
End Sub

Private Sub mnuNegativenet_Click()
    If allowedAccess("frmRptNetPayListNegative") = False Then Exit Sub
    frmRptNetPayListNegative.Show
End Sub

Private Sub mnuNetPay_Click()
    If allowedAccess("frmRptNetPayList") = False Then Exit Sub
    frmRptNetPayList.Show
End Sub

Private Sub mnuNetPayCap_Click()
    If allowedAccess(frmMDNetPayCap.Name) = False Then Exit Sub
    frmMDNetPayCap.Show
    frmMDNetPayCap.ZOrder
End Sub

Private Sub mnuOBEntry_Click()
    If allowedAccess("frmLOBBusiness") = False Then Exit Sub
    frmLOBBusiness.Show
    frmLOBBusiness.ZOrder
End Sub

Private Sub mnuOtherBonusType_Click()
    If allowedAccess("frmMDOtherBonus") = False Then Exit Sub
    With frmMDOtherBonus
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuOtherDeduct_Click()
    If allowedAccess("frmMDOtherDed") = False Then Exit Sub
    With frmMDOtherDed
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnuOtherDeductions_Click()
    If allowedAccess("frmAdOtherDeductions") = False Then Exit Sub
    frmAdOtherDeductions.Show
    frmAdOtherDeductions.ZOrder
End Sub

Private Sub mnuOtherDeductRep_Click()
    If allowedAccess("frmRptOtherDeductions") = False Then Exit Sub
  frmRptOtherDeductions.Show
  frmRptOtherDeductions.ZOrder
End Sub

Private Sub mnuOtherEarning_Click()
    If allowedAccess("frmMDOtherEarn") = False Then Exit Sub
    With frmMDOtherEarn
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuOtherEarnings_Click()
    If allowedAccess("frmAdOtherEarnings") = False Then Exit Sub
    frmAdOtherEarnings.Show
    frmAdOtherEarnings.ZOrder
End Sub

Private Sub mnuOvertime_Click()
    If allowedAccess("frmADOvertime") = False Then Exit Sub
    frmADOvertime.Show
    frmADOvertime.ZOrder
End Sub

Private Sub mnuParameter_Click()
    If allowedAccess("frmParameter") = False Then Exit Sub
    frmParameter.Show vbModal
End Sub

Private Sub mnuPayFrequency_Click()
    If allowedAccess("frmMDPayFreq") = False Then Exit Sub
    frmMDPayFreq.Show
    frmMDPayFreq.ZOrder
End Sub

Private Sub mnuPayrollPeriod_Click()
    If allowedAccess("frmMDPayrollPeriod") = False Then Exit Sub
'    frmMDPeriod.Show
'    frmMDPeriod.ZOrder
    frmMDPayrollPeriod.Show
    frmMDPayrollPeriod.ZOrder
End Sub

Private Sub mnuPayrollReg_Click()
    If allowedAccess("frmRPTPayrollRegister") = False Then Exit Sub
    frmRPTPayrollRegister.Show
    frmRPTPayrollRegister.ZOrder
End Sub

Private Sub mnuPayrollSumm_Click()
    If allowedAccess("frmPPSummary") = False Then Exit Sub
    frmPPSummary.Show
    frmPPSummary.ZOrder
End Sub

Private Sub mnuPayslip_Click()
    If allowedAccess("frmRptPayslips") = False Then Exit Sub
    frmRptPayslips.Show
    frmRptPayslips.ZOrder
End Sub

Private Sub mnuPHICTable_Click()
    If allowedAccess("frmMDPhilhealth") = False Then Exit Sub
    With frmMDPhilhealth
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuPhilHealthContribution_Click()
    If allowedAccess("frmRptPhilHealthContribution") = False Then Exit Sub
    frmRptPhilHealthContribution.Show
    frmRptPhilHealthContribution.ZOrder
End Sub

Private Sub mnuRateType_Click()
    If allowedAccess("frmMDRate") = False Then Exit Sub
    With frmMDRate
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuReceivables_Click()
    If allowedAccess("frmAdReceivables") = False Then Exit Sub
    frmAdReceivables.Show
    frmAdReceivables.ZOrder
End Sub

Private Sub mnuShiftSched_Click()
    If allowedAccess("frmMDShift") = False Then Exit Sub
    frmMDShift.Show
    frmMDShift.ZOrder
End Sub

Private Sub mnuSSSContributions_Click()
    If allowedAccess("frmRptSSSContribution") = False Then Exit Sub
    frmRptSSSContribution.Show
    frmRptSSSContribution.ZOrder
End Sub

Private Sub mnuSSSEPF_Click()
    If allowedAccess("frmRptSSSAlphaList") = False Then Exit Sub
  frmRptSSSAlphaList.Show
  frmRptSSSAlphaList.ZOrder
End Sub

Private Sub mnuSSSTable_Click()
    If allowedAccess("frmMDSSS") = False Then Exit Sub
    With frmMDSSS
        .Show
        .SetFocus
    End With
End Sub

Private Sub mnuTITO_Click()
    If allowedAccess("frmPPImportTito") = False Then Exit Sub
  frmPPImportTito.Show
  frmPPImportTito.ZOrder
End Sub

Private Sub mnuUserGroups_Click()
    frmMDUsergroups.Show
    frmMDUsergroups.ZOrder
End Sub

Private Sub mnuUsers_Click()
    If allowedAccess("frmMDUsers") = False Then Exit Sub
    frmMDUsers.Show
    frmMDUsers.ZOrder
End Sub

Private Sub mnuVoucher_Click()
    If allowedAccess("frmAdVouchers") = False Then Exit Sub
    frmAdVouchers.Show
    frmAdVouchers.ZOrder
End Sub

Private Sub mnuVoucherSlips_Click()
    If allowedAccess("frmRptVoucherSlips") = False Then Exit Sub
    frmRptVoucherSlips.Show
    frmRptVoucherSlips.ZOrder
End Sub

Private Sub mnuVoucherSummary_Click()
    If allowedAccess("frmRptVoucherSummary") = False Then Exit Sub
    frmRptVoucherSummary.Show
    frmRptVoucherSummary.ZOrder
End Sub

Private Sub mnuWithholdingTaxContribution_Click()
    If allowedAccess("frmRptWithHoldingsTaxContribution") = False Then Exit Sub
    frmRptWithHoldingsTaxContribution.Show
    frmRptWithHoldingsTaxContribution.ZOrder
End Sub

Private Sub mnuWTaxTable_Click()
    If allowedAccess("frmMDWithhold") = False Then Exit Sub
    With frmMDWithhold
        .Show
        .SetFocus
    End With
End Sub

Private Sub mSection_Click()
    If allowedAccess("frmMDSection") = False Then Exit Sub
    frmMDSection.Show
    frmMDSection.ZOrder
End Sub

Private Sub Municipal_Click()
    If allowedAccess("frmMDMunicipal") = False Then Exit Sub
    frmMDMunicipal.Show
    frmMDMunicipal.ZOrder
End Sub

Private Sub province_Click()
    If allowedAccess("frmMDProvince") = False Then Exit Sub
    frmMDProvince.Show
    frmMDProvince.ZOrder
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
      
        Case "mnuPayrollPeriod":    frmMDPayrollPeriod.Show
                                    frmMDPayrollPeriod.ZOrder
        Case "mnuLoanApplications": frmADLoans.Show
                                    frmADLoans.ZOrder
        Case "mnuOvertime":         frmADOvertime.Show
                                    frmADOvertime.ZOrder
        Case "mnuOtherEarnings":    frmAdOtherEarnings.Show
                                    frmAdOtherEarnings.ZOrder
        Case "mnuOtherDeductions":  frmAdOtherDeductions.Show
                                    frmAdOtherDeductions.ZOrder
        Case "mnuLeaves":           frmLOBLeave.Show
                                    frmLOBLeave.ZOrder
        Case "mnuGenLoan":          frmPPGenLoanDed.Show
                                    frmPPGenLoanDed.ZOrder
        Case "mnuDtrSummary":       frmPPDTR.Show
                                    frmPPDTR.ZOrder
        Case "mnuGenPayroll":       frmPPGeneratePayroll.Show vbModal
        
        Case "mnuFinalizePayroll":  frmPPFinalizePayroll.Show vbModal
        
    End Select
    
End Sub

Private Function allowedAccess(ByVal formName As String) As Boolean
'  If formName = "frmPPDTR" Then
    allowedAccess = True
'  Else
'    allowedAccess = False
'  End If
End Function


