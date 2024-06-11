VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmRptAlphaList 
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   495
      TabIndex        =   2
      Top             =   45
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   609
      Caption         =   "Export Alphalist"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   4210752
   End
   Begin VB.Frame fraParmtr 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   8820
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   4395
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Meal Deduction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   21
         Left            =   180
         TabIndex        =   32
         Tag             =   "x1.mealallow"
         Top             =   6570
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   23
         Left            =   180
         TabIndex        =   29
         Tag             =   "x1.isactive"
         Top             =   7125
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Date hired"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   22
         Left            =   180
         TabIndex        =   28
         Tag             =   "x1.datehired"
         Top             =   6840
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Fixed Earnings"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   20
         Left            =   180
         TabIndex        =   27
         Tag             =   "x1.fixedearnings"
         Top             =   6300
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Tax bracket"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   19
         Left            =   180
         TabIndex        =   22
         Tag             =   "x7.description"
         Top             =   6030
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Employment Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   18
         Left            =   180
         TabIndex        =   21
         Tag             =   "x6.empstatname"
         Top             =   5760
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Rate type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   17
         Left            =   180
         TabIndex        =   20
         Tag             =   "x5.ratetypename"
         Top             =   5490
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Hourly rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   16
         Left            =   180
         TabIndex        =   19
         Tag             =   "x1.hourly_rate"
         Top             =   5220
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Daily rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   15
         Left            =   180
         TabIndex        =   18
         Tag             =   "x1.daily_rate"
         Top             =   4950
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Monthly rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   14
         Left            =   180
         TabIndex        =   17
         Tag             =   "x1.monthly_rate"
         Top             =   4695
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Account no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   13
         Left            =   180
         TabIndex        =   16
         Tag             =   "x1.bankacctno"
         Top             =   4425
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "T.I.N. no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   12
         Left            =   180
         TabIndex        =   15
         Tag             =   "x1.tinno"
         Top             =   4140
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Pag-IBIG"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   11
         Left            =   180
         TabIndex        =   14
         Tag             =   "x1.hdmfno"
         Top             =   3870
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Philhealth no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   13
         Tag             =   "x1.philhno"
         Top             =   3600
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "SSS no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   12
         Tag             =   "x1.sssno"
         Top             =   3330
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Tel/Mobile no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   8
         Left            =   180
         TabIndex        =   11
         Tag             =   "concat(x1.telno,"";"",x1.mobileno)"
         Top             =   3075
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   7
         Left            =   180
         TabIndex        =   10
         Tag             =   "x1.street"
         Top             =   2805
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Birth date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   6
         Left            =   180
         TabIndex        =   9
         Tag             =   "date(x1.birthdate) birthdate"
         Top             =   2535
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Civil status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Tag             =   "x1.civilstatus"
         Top             =   2265
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Tag             =   "x1.gender"
         Top             =   2010
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Cost Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   6
         Tag             =   "x3.costcenter"
         Top             =   1740
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Tag             =   "x2.division"
         Top             =   1470
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Job title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Tag             =   "x4.description jobtitle"
         Top             =   1200
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Tag             =   "x1.dummycode"
         Top             =   915
         Width           =   2340
      End
      Begin OsenXPCntrl.OsenXPButton cmdCheckAll 
         Height          =   360
         Left            =   180
         TabIndex        =   23
         Top             =   90
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         BTYPE           =   5
         TX              =   "&Check All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16185592
         BCOLO           =   16185592
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmRptAlphaList.frx":0000
         PICN            =   "frmRptAlphaList.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdUncheckAll 
         Height          =   360
         Left            =   1800
         TabIndex        =   24
         Top             =   90
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   635
         BTYPE           =   5
         TX              =   "&Uncheck All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16185592
         BCOLO           =   16185592
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmRptAlphaList.frx":05B6
         PICN            =   "frmRptAlphaList.frx":05D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin lvButton.lvButtons_H cmdExport 
         Height          =   375
         Left            =   2310
         TabIndex        =   25
         Top             =   8280
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         Caption         =   "&Save as Excel File"
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
         Image           =   "frmRptAlphaList.frx":0B6C
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdView 
         Height          =   375
         Left            =   195
         TabIndex        =   26
         Top             =   8280
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         Caption         =   "&View"
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
         Image           =   "frmRptAlphaList.frx":678E
         cBack           =   14737632
      End
      Begin TrueOleDBList80.TDBCombo tdbName 
         Height          =   345
         Left            =   825
         TabIndex        =   30
         Tag             =   "Municipal"
         Top             =   525
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   609
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   609
         _GAPHEIGHT      =   53
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Name"
         Columns(0).DataField=   "nameoption"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits.Count    =   1
         Appearance      =   0
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   -1  'True
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   -1  'True
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   3
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   2
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         MaxComboItems   =   3
         AddItemSeparator=   ";"
         _PropDict       =   $"frmRptAlphaList.frx":7468
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(36)  =   "Named:id=33:Normal"
         _StyleDefs(37)  =   ":id=33,.parent=0"
         _StyleDefs(38)  =   "Named:id=34:Heading"
         _StyleDefs(39)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(40)  =   ":id=34,.wraptext=-1"
         _StyleDefs(41)  =   "Named:id=35:Footing"
         _StyleDefs(42)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(43)  =   "Named:id=36:Selected"
         _StyleDefs(44)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(45)  =   "Named:id=37:Caption"
         _StyleDefs(46)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(47)  =   "Named:id=38:HighlightRow"
         _StyleDefs(48)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=39:EvenRow"
         _StyleDefs(50)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(51)  =   "Named:id=40:OddRow"
         _StyleDefs(52)  =   ":id=40,.parent=33"
         _StyleDefs(53)  =   "Named:id=41:RecordSelector"
         _StyleDefs(54)  =   ":id=41,.parent=34"
         _StyleDefs(55)  =   "Named:id=42:FilterBar"
         _StyleDefs(56)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   195
         TabIndex        =   31
         Top             =   585
         Width           =   1035
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   8085
      Left            =   6660
      TabIndex        =   0
      Top             =   525
      Width           =   3195
      _cx             =   5636
      _cy             =   14261
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   128
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8421504
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.xls"
   End
End
Attribute VB_Name = "frmRptAlphaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sheet%

Dim rsNameOption  As ADODB.Recordset

Private Sub cmdCheckAll_Click()
    
    Dim i               As Integer
    
    For i = 0 To chkField.UBound
        chkField(i).Value = 1
    Next
    
End Sub

Private Sub cmdExport_Click()

     On Error GoTo ErrHndlr

    dlg.FileName = ""
    dlg.ShowSave
    If Len(dlg.FileName) = 0 Then Exit Sub

    MousePointer = MousePointerConstants.vbHourglass
    fg.SaveGrid dlg.FileName, flexFileExcel
    MousePointer = MousePointerConstants.vbDefault
    
    'Caption = dlg.FileName

    'sheet = 0

    
ErrHndlr:
End Sub

Private Sub cmdUncheckAll_Click()
    
    Dim i               As Integer
    
    For i = 0 To chkField.UBound
        chkField(i).Value = 0
    Next
    
End Sub

Private Sub cmdView_Click()
    
    Dim i                   As Integer
    Dim mCol                As Integer
    Dim mSkipCol            As Integer
    Dim mRow                As Integer
    Dim mRow2               As Integer
    
    Dim mQuery              As String
    Dim mTitle              As String
    
    Dim rsQuery             As ADODB.Recordset
    
    Select Case tdbName.Text
      Case "LAST NAME | FIRST NAME | MIDDLE NAME":
            mQuery = "x1.lastname,x1.firstname,x1.middlename "
            mSkipCol = 3
      Case "FIRST NAME MIDDLE NAME LAST NAME":
            mQuery = "concat(x1.firstname,' ',x1.middlename,' ',x1.lastname) employeename "
            mSkipCol = 1
      Case "LAST NAME, FIRST NAME MIDDLE NAME":
            mQuery = "concat(x1.lastname,', ',x1.firstname,' ',x1.middlename) employeename "
            mSkipCol = 1
    End Select
    
    For i = 0 To chkField.UBound
        If chkField(i).Value = 1 Then
            If Trim(mQuery) = "" Then
                mTitle = "'" & chkField(i).Caption & "'"
                mQuery = chkField(i).Tag
            Else
                mTitle = mTitle & "," & "'" & chkField(i).Caption & "'"
                mQuery = mQuery & "," & chkField(i).Tag
            End If
        End If
    Next
    
    If Trim(mQuery) = "" Then
        
        MsgBox "Please select a column to display.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    Dim mstr As String
    Debug.Print
    mstr = "select " & mQuery & ",x9.rice,x9.laundry,x9.clothing,x9.medical from employee x1 " & _
                    "left outer join division x2 on x1.divisioncode = x2.divisioncode " & _
                    "left outer join costcenter x3 on x1.costcentercode = x3.costcentercode " & _
                    "left outer join jobtitle x4 on x1.jobtitlecode = x4.jobtitlecode " & _
                    "left outer join ratetypes x5 on x1.ratetypecode = x5.ratetypecode " & _
                    "left outer join employmentstatus x6 on x1.empstatcode = x6.empstatcode " & _
                    "left outer join wt x7 on x1.wtcode = x7.wtcode " & _
                    "left outer join payfrequency x8 on x1.payfreqcode = x8.payfreqcode " & _
                    "left outer join (SELECT EMPLOYEECODE, SUM(RICE) RICE,SUM(LAUNDRY) LAUNDRY,SUM(CLOTHING) CLOTHING, SUM(MEDICAL) MEDICAL FROM " & _
                                      "(SELECT EMPLOYEECODE,NONTAXALLOW_AMT RICE,0 LAUNDRY,0 CLOTHING,0 MEDICAL FROM employee_nontaxallow WHERE NONTAXALLOW_ID = 1 " & _
                                      "Union All " & _
                                      "SELECT EMPLOYEECODE,0 RICE,NONTAXALLOW_AMT LAUNDRY,0 CLOTHING,0 MEDICAL FROM employee_nontaxallow WHERE NONTAXALLOW_ID = 2 " & _
                                      "Union All " & _
                                      "SELECT EMPLOYEECODE,0 RICE,0 LAUNDRY,NONTAXALLOW_AMT CLOTHING,0 MEDICAL FROM employee_nontaxallow WHERE NONTAXALLOW_ID = 3 " & _
                                      "Union All " & _
                                      "SELECT EMPLOYEECODE,0 RICE,0 LAUNDRY,0 CLOTHING,NONTAXALLOW_AMT MEDICAL FROM employee_nontaxallow WHERE NONTAXALLOW_ID = 4) S1 " & _
                                      "GROUP BY EMPLOYEECODE) x9 on x1.employeecode=x9.employeecode "
        NetOpen rsQuery, mstr
     Debug.Print mstr
    Set fg.DataSource = rsQuery
    
    If rsQuery.RecordCount > 0 Then
        
        fg.FrozenRows = 1
        fg.Rows = fg.Rows + 1
        mRow = fg.Rows - 1
        mRow2 = fg.Rows - 2
        
        Do While mRow >= 1
        
            mCol = 0
            
            For mCol = 0 To fg.Cols - 1
                fg.TextMatrix(mRow, mCol) = fg.TextMatrix(mRow2, mCol)
            Next
            
            mRow = mRow - 1
            mRow2 = mRow2 - 1
            
        Loop
        
'        Select Case tdbName.Text
'          Case "LAST NAME | FIRST NAME | MIDDLE NAME":
'                fg.TextMatrix(0, 0) = "Last Name"
'                fg.TextMatrix(0, 1) = "First Name"
'                fg.TextMatrix(0, 2) = "Middle Name"
'          Case "FIRST NAME MIDDLE NAME LAST NAME":
'                fg.TextMatrix(0, 0) = "Employee Name"
'          Case "LAST NAME, FIRST NAME MIDDLE NAME":
'                fg.TextMatrix(0, 0) = "Employee Name"
'        End Select

'If tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 0 Then
'                    fg.TextMatrix(0, mCol) = "Last Name"
'            ElseIf tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 1 Then
'                    fg.TextMatrix(0, mCol) = "First Name"
'            ElseIf tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 2 Then
'                    fg.TextMatrix(0, mCol) = "Middle Name"
'            ElseIf tdbName.Text = "FIRST NAME MIDDLE NAME LAST NAME" And mCol = 0 Then
'                    fg.TextMatrix(0, mCol) = "Employee Name"
'            ElseIf tdbName.Text = "LAST NAME, FIRST NAME MIDDLE NAME" And mCol = 0 Then
'                    fg.TextMatrix(0, mCol) = "Employee Name"
'            Else
        
        mCol = 0
        
        For mCol = 0 To mSkipCol - 1
              If fg.ColWidth(mCol) < 2000 Then fg.ColWidth(mCol) = 2000
              fg.Cell(flexcpBackColor, 0, mCol, 0, mCol) = &H808080
              fg.Cell(flexcpFontBold, 0, mCol, 0, mCol) = True
              fg.Cell(flexcpForeColor, 0, mCol, 0, mCol) = vbWhite
              If tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 0 Then
                      fg.TextMatrix(0, mCol) = "Last Name"
              ElseIf tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 1 Then
                      fg.TextMatrix(0, mCol) = "First Name"
              ElseIf tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME" And mCol = 2 Then
                      fg.TextMatrix(0, mCol) = "Middle Name"
              ElseIf tdbName.Text = "FIRST NAME MIDDLE NAME LAST NAME" And mCol = 0 Then
                      fg.TextMatrix(0, mCol) = "Employee Name"
              ElseIf tdbName.Text = "LAST NAME, FIRST NAME MIDDLE NAME" And mCol = 0 Then
                      fg.TextMatrix(0, mCol) = "Employee Name"
              End If
        Next
        
        i = 0

        mCol = mSkipCol
        For i = 0 To chkField.UBound
            If chkField(i).Value = 1 Then
              If fg.ColWidth(mCol) < 2000 Then fg.ColWidth(mCol) = 2000
              fg.Cell(flexcpBackColor, 0, mCol, 0, mCol) = &H808080
              fg.Cell(flexcpFontBold, 0, mCol, 0, mCol) = True
              fg.Cell(flexcpForeColor, 0, mCol, 0, mCol) = vbWhite
              fg.TextMatrix(0, mCol) = chkField(i).Caption
              mCol = mCol + 1
            End If
        Next
        
        If fg.ColWidth(mCol) < 2000 Then fg.ColWidth(mCol) = 2000
        fg.Cell(flexcpBackColor, 0, mCol, 0, mCol) = &H808080
        fg.Cell(flexcpFontBold, 0, mCol, 0, mCol) = True
        fg.Cell(flexcpForeColor, 0, mCol, 0, mCol) = vbWhite
        fg.ColFormat(mCol) = "#,##0.00"
        fg.TextMatrix(0, mCol) = "Rice"

        If fg.ColWidth(mCol + 1) < 2000 Then fg.ColWidth(mCol + 1) = 2000
        fg.Cell(flexcpBackColor, 0, mCol + 1, 0, mCol + 1) = &H808080
        fg.Cell(flexcpFontBold, 0, mCol + 1, 0, mCol + 1) = True
        fg.Cell(flexcpForeColor, 0, mCol + 1, 0, mCol + 1) = vbWhite
        fg.ColFormat(mCol + 1) = "#,##0.00"
        fg.TextMatrix(0, mCol + 1) = "Laundry"

        If fg.ColWidth(mCol + 2) < 2000 Then fg.ColWidth(mCol + 2) = 2000
        fg.Cell(flexcpBackColor, 0, mCol + 2, 0, mCol + 2) = &H808080
        fg.Cell(flexcpFontBold, 0, mCol + 2, 0, mCol + 2) = True
        fg.Cell(flexcpForeColor, 0, mCol + 2, 0, mCol + 2) = vbWhite
        fg.ColFormat(mCol + 2) = "#,##0.00"
        fg.TextMatrix(0, mCol + 2) = "Clothing"
        
        If fg.ColWidth(mCol + 3) < 2000 Then fg.ColWidth(mCol + 3) = 2000
        fg.Cell(flexcpBackColor, 0, mCol + 3, 0, mCol + 3) = &H808080
        fg.Cell(flexcpFontBold, 0, mCol + 3, 0, mCol + 3) = True
        fg.Cell(flexcpForeColor, 0, mCol + 3, 0, mCol + 3) = vbWhite
        fg.ColFormat(mCol + 3) = "#,##0.00"
        fg.TextMatrix(0, mCol + 3) = "Medical"

    End If
    
End Sub

Private Sub Form_Load()

    Dim i     As Integer
    
    Add_MDIButton Me.Name, TitleBar.Caption
    
    dlg.DefaultExt = "xls"
    dlg.Filter = "Excel 97 (*.xls)|*.xls"
    
    Set rsNameOption = New ADODB.Recordset
    
    With rsNameOption
      .Fields.Append "nameoption", adVarChar, 100
      .Open
      For i = 1 To 3
        .AddNew
        Select Case i
          Case 1: .Fields("nameoption") = "LAST NAME | FIRST NAME | MIDDLE NAME"
          Case 2: .Fields("nameoption") = "FIRST NAME MIDDLE NAME LAST NAME"
          Case 3: .Fields("nameoption") = "LAST NAME, FIRST NAME MIDDLE NAME"
        End Select
        .Update
      Next
    End With

    Set tdbName.RowSource = rsNameOption
    tdbName.Text = "LAST NAME | FIRST NAME | MIDDLE NAME"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fraParmtr
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Height = Me.ScaleHeight - .Top
    End With

    With fg
        .Top = TitleBar.Top + TitleBar.Height
        .Left = fraParmtr.Left + fraParmtr.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top
    End With
    
    With cmdView
        .Top = fraParmtr.Height - .Height
    End With
   
    With cmdExport
        .Top = fraParmtr.Height - .Height
    End With
   
End Sub

