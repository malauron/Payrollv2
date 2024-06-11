VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDPeriod 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8220
   ClientLeft      =   6990
   ClientTop       =   3300
   ClientWidth     =   12825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMDPeriod.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab tabPayrollPeriod 
      Height          =   7980
      Left            =   30
      TabIndex        =   18
      Top             =   615
      Width           =   12780
      _cx             =   22542
      _cy             =   14076
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16185592
      ForeColor       =   -2147483630
      FrontTabColor   =   16185592
      BackTabColor    =   16185592
      TabOutlineColor =   12632256
      FrontTabForeColor=   0
      Caption         =   "Maintain Payroll Periods|View Payroll Periods"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   4
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic SizerMAType 
         Height          =   7665
         Left            =   15
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   12750
         _cx             =   22490
         _cy             =   13520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16185592
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame FraInfo 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   7440
            Left            =   195
            TabIndex        =   33
            Top             =   105
            Width           =   11385
            Begin VB.Frame Frame5 
               BackColor       =   &H00F6F8F8&
               Caption         =   "Loan Deductions"
               Height          =   4560
               Left            =   3165
               TabIndex        =   55
               Top             =   2610
               Width           =   5895
               Begin TrueOleDBGrid80.TDBGrid tdgLeaveLimit 
                  Height          =   4260
                  Left            =   60
                  TabIndex        =   56
                  Top             =   225
                  Width           =   5760
                  _ExtentX        =   10160
                  _ExtentY        =   7514
                  _LayoutType     =   4
                  _RowHeight      =   16
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "Loan Types"
                  Columns(0).DataField=   "loantypesname"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   4
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   "allow"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).DataField=   ""
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   3
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).MarqueeStyle=   4
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0).RecordSelectors=   0   'False
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0)._GSX_SAVERECORDSELECTORS=   0
                  Splits(0).AllowColSelect=   0   'False
                  Splits(0).AlternatingRowStyle=   -1  'True
                  Splits(0).DividerColor=   16777215
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=3"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=8493"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8414"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=503"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerStyle=0"
                  Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=450"
                  Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
                  Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(14)=   "Column(1)._HeadDivider=0"
                  Splits(0)._ColumnProps(15)=   "Column(2).Width=79"
                  Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
                  Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   0
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
                  PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   2
                  DefColWidth     =   0
                  HeadLines       =   2
                  FootLines       =   1
                  RowDividerStyle =   0
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTipsWidth   =   0
                  DeadAreaBackColor=   12632256
                  RowDividerColor =   13160660
                  RowSubDividerColor=   16777215
                  DirectionAfterEnter=   1
                  DirectionAfterTab=   0
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=900"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Arial"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HF6F8F8&"
                  _StyleDefs(11)  =   ":id=2,.fgcolor=&H400000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
                  _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
                  _StyleDefs(13)  =   ":id=2,.fontname=Arial"
                  _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.borderColor=&H808080&"
                  _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HF6F8F8&,.appearance=1"
                  _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&"
                  _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFFF0EA&"
                  _StyleDefs(19)  =   ":id=8,.fgcolor=&H0&"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFFFF&"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HF6F8F8&"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&H808080&,.fgcolor=&H80FFFF&"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&HFFF0EA&"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13,.alignment=2,.locked=0"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
                  _StyleDefs(48)  =   "Named:id=33:Normal"
                  _StyleDefs(49)  =   ":id=33,.parent=0"
                  _StyleDefs(50)  =   "Named:id=34:Heading"
                  _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(52)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(53)  =   "Named:id=35:Footing"
                  _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(55)  =   "Named:id=36:Selected"
                  _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(57)  =   "Named:id=37:Caption"
                  _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(59)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
                  _StyleDefs(61)  =   "Named:id=39:EvenRow"
                  _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(63)  =   "Named:id=40:OddRow"
                  _StyleDefs(64)  =   ":id=40,.parent=33"
                  _StyleDefs(65)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(66)  =   ":id=41,.parent=34"
                  _StyleDefs(67)  =   "Named:id=42:FilterBar"
                  _StyleDefs(68)  =   ":id=42,.parent=33"
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00F6F8F8&
               Caption         =   "HDMF Contribution"
               Height          =   1140
               Left            =   45
               TabIndex        =   54
               Top             =   6030
               Width           =   3000
               Begin VB.CheckBox chkHDMFDaily 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Daily Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   14
                  Top             =   315
                  Width           =   2520
               End
               Begin VB.CheckBox chkHDMFMonthly 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Monthly Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   15
                  Top             =   615
                  Width           =   2520
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00F6F8F8&
               Caption         =   "Withholding Tax"
               Height          =   1140
               Left            =   45
               TabIndex        =   53
               Top             =   4890
               Width           =   3000
               Begin VB.CheckBox chkTaxMonthly 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Monthly Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   13
                  Top             =   615
                  Width           =   2520
               End
               Begin VB.CheckBox chkTaxDaily 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Daily Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   12
                  Top             =   315
                  Width           =   2520
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00F6F8F8&
               Caption         =   "PhilHealth Contribution"
               Height          =   1140
               Left            =   45
               TabIndex        =   52
               Top             =   3750
               Width           =   3000
               Begin VB.CheckBox chkPHDaily 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Daily Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   10
                  Top             =   315
                  Width           =   2520
               End
               Begin VB.CheckBox chkPHMonthly 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Monthly Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   11
                  Top             =   615
                  Width           =   2520
               End
            End
            Begin VB.Frame fraSSS 
               BackColor       =   &H00F6F8F8&
               Caption         =   "SSS Contribution"
               Height          =   1140
               Left            =   45
               TabIndex        =   51
               Top             =   2610
               Width           =   3000
               Begin VB.CheckBox chkSSSMonthly 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Monthly Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   9
                  Top             =   615
                  Width           =   2520
               End
               Begin VB.CheckBox chkSSSDaily 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Daily Rate Employees"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Left            =   255
                  TabIndex        =   8
                  Top             =   315
                  Width           =   2520
               End
            End
            Begin OsenXPCntrl.OsenXPButton cmdPayFreqShow 
               Height          =   300
               Left            =   7515
               TabIndex        =   34
               Top             =   75
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   529
               BTYPE           =   3
               TX              =   "..."
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
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   16777215
               BCOLO           =   16777215
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   12632256
               MPTR            =   0
               MICON           =   "frmMDPeriod.frx":6852
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin TDBText6Ctl.TDBText txtPayPeriod 
               Height          =   300
               Left            =   1935
               TabIndex        =   4
               Top             =   1035
               Width           =   2010
               _Version        =   65536
               _ExtentX        =   3545
               _ExtentY        =   529
               Caption         =   "frmMDPeriod.frx":686E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":68DA
               Key             =   "frmMDPeriod.frx":68F8
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   -1
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   0
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   2
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate txtTo 
               Height          =   300
               Left            =   1935
               TabIndex        =   7
               Top             =   2025
               Width           =   2010
               _Version        =   65536
               _ExtentX        =   3545
               _ExtentY        =   529
               Calendar        =   "frmMDPeriod.frx":693C
               Caption         =   "frmMDPeriod.frx":6A42
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":6AA8
               Keys            =   "frmMDPeriod.frx":6AC6
               Spin            =   "frmMDPeriod.frx":6B24
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "mm/dd/yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "mm/dd/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "01/16/2008"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   39463
               CenturyMode     =   0
            End
            Begin TrueOleDBList80.TDBCombo tdbMonth 
               Height          =   345
               Left            =   1965
               TabIndex        =   2
               Tag             =   "Municipal"
               Top             =   405
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   609
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               _DropdownWidth  =   0
               _EDITHEIGHT     =   609
               _GAPHEIGHT      =   53
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Code"
               Columns(0).DataField=   "provcode"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Descritpion"
               Columns(1).DataField=   "provdesc"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits.Count    =   1
               Appearance      =   0
               BorderStyle     =   1
               ComboStyle      =   0
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
               _PropDict       =   $"frmMDPeriod.frx":6B4C
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
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Named:id=33:Normal"
               _StyleDefs(41)  =   ":id=33,.parent=0"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(44)  =   ":id=34,.wraptext=-1"
               _StyleDefs(45)  =   "Named:id=35:Footing"
               _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(47)  =   "Named:id=36:Selected"
               _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(49)  =   "Named:id=37:Caption"
               _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(51)  =   "Named:id=38:HighlightRow"
               _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=39:EvenRow"
               _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(55)  =   "Named:id=40:OddRow"
               _StyleDefs(56)  =   ":id=40,.parent=33"
               _StyleDefs(57)  =   "Named:id=41:RecordSelector"
               _StyleDefs(58)  =   ":id=41,.parent=34"
               _StyleDefs(59)  =   "Named:id=42:FilterBar"
               _StyleDefs(60)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBList80.TDBCombo tdbClassification 
               Height          =   345
               Left            =   5505
               TabIndex        =   3
               Tag             =   "Municipal"
               Top             =   405
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   609
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               _DropdownWidth  =   0
               _EDITHEIGHT     =   609
               _GAPHEIGHT      =   53
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Code"
               Columns(0).DataField=   "provcode"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Descritpion"
               Columns(1).DataField=   "provdesc"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits.Count    =   1
               Appearance      =   0
               BorderStyle     =   1
               ComboStyle      =   0
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
               _PropDict       =   $"frmMDPeriod.frx":6BF6
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
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Named:id=33:Normal"
               _StyleDefs(41)  =   ":id=33,.parent=0"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(44)  =   ":id=34,.wraptext=-1"
               _StyleDefs(45)  =   "Named:id=35:Footing"
               _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(47)  =   "Named:id=36:Selected"
               _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(49)  =   "Named:id=37:Caption"
               _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(51)  =   "Named:id=38:HighlightRow"
               _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=39:EvenRow"
               _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(55)  =   "Named:id=40:OddRow"
               _StyleDefs(56)  =   ":id=40,.parent=33"
               _StyleDefs(57)  =   "Named:id=41:RecordSelector"
               _StyleDefs(58)  =   ":id=41,.parent=34"
               _StyleDefs(59)  =   "Named:id=42:FilterBar"
               _StyleDefs(60)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBList80.TDBCombo tdbPayFrequency 
               Height          =   345
               Left            =   5535
               TabIndex        =   1
               Tag             =   "Municipal"
               Top             =   30
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   609
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               _DropdownWidth  =   0
               _EDITHEIGHT     =   609
               _GAPHEIGHT      =   53
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Code"
               Columns(0).DataField=   "provcode"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Descritpion"
               Columns(1).DataField=   "provdesc"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).AllowRowSizing=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
               Splits.Count    =   1
               Appearance      =   0
               BorderStyle     =   1
               ComboStyle      =   0
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
               MaxComboItems   =   10
               AddItemSeparator=   ";"
               _PropDict       =   $"frmMDPeriod.frx":6CA0
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
               _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Named:id=33:Normal"
               _StyleDefs(41)  =   ":id=33,.parent=0"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(44)  =   ":id=34,.wraptext=-1"
               _StyleDefs(45)  =   "Named:id=35:Footing"
               _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(47)  =   "Named:id=36:Selected"
               _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(49)  =   "Named:id=37:Caption"
               _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(51)  =   "Named:id=38:HighlightRow"
               _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=39:EvenRow"
               _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(55)  =   "Named:id=40:OddRow"
               _StyleDefs(56)  =   ":id=40,.parent=33"
               _StyleDefs(57)  =   "Named:id=41:RecordSelector"
               _StyleDefs(58)  =   ":id=41,.parent=34"
               _StyleDefs(59)  =   "Named:id=42:FilterBar"
               _StyleDefs(60)  =   ":id=42,.parent=33"
            End
            Begin TDBText6Ctl.TDBText txtDescription 
               Height          =   300
               Left            =   1935
               TabIndex        =   5
               Top             =   1365
               Width           =   3600
               _Version        =   65536
               _ExtentX        =   6350
               _ExtentY        =   529
               Caption         =   "frmMDPeriod.frx":6D4A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":6DB6
               Key             =   "frmMDPeriod.frx":6DD4
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   0
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   2
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBNumber6Ctl.TDBNumber txtPayyear 
               Height          =   345
               Left            =   1965
               TabIndex        =   0
               Top             =   30
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   609
               Calculator      =   "frmMDPeriod.frx":6E18
               Caption         =   "frmMDPeriod.frx":6E38
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":6E9E
               Keys            =   "frmMDPeriod.frx":6EBC
               Spin            =   "frmMDPeriod.frx":6F06
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#####"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999
               MinValue        =   -99999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   1245189
               Value           =   0
               MaxValueVT      =   1145962501
               MinValueVT      =   1414463493
            End
            Begin TDBDate6Ctl.TDBDate txtFrom 
               Height          =   300
               Left            =   1935
               TabIndex        =   6
               Top             =   1695
               Width           =   2010
               _Version        =   65536
               _ExtentX        =   3545
               _ExtentY        =   529
               Calendar        =   "frmMDPeriod.frx":6F2E
               Caption         =   "frmMDPeriod.frx":7034
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":709A
               Keys            =   "frmMDPeriod.frx":70B8
               Spin            =   "frmMDPeriod.frx":7116
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "mm/dd/yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "mm/dd/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "01/16/2008"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   39463
               CenturyMode     =   0
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000010&
               X1              =   60
               X2              =   9120
               Y1              =   2460
               Y2              =   2460
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               Height          =   255
               Left            =   270
               TabIndex        =   42
               Top             =   1395
               Width           =   1560
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000010&
               X1              =   60
               X2              =   9105
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Period: From"
               Height          =   255
               Left            =   -15
               TabIndex        =   41
               Top             =   1740
               Width           =   1845
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Period Code"
               Height          =   255
               Left            =   270
               TabIndex        =   40
               Top             =   1065
               Width           =   1560
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "To"
               Height          =   255
               Left            =   600
               TabIndex        =   39
               Top             =   2055
               Width           =   1230
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Payroll Year"
               Height          =   255
               Left            =   330
               TabIndex        =   38
               Top             =   105
               Width           =   1560
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Month"
               Height          =   255
               Index           =   0
               Left            =   300
               TabIndex        =   37
               Top             =   435
               Width           =   1560
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Week"
               Height          =   255
               Left            =   4185
               TabIndex        =   36
               Top             =   450
               Width           =   1200
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pay Frequency"
               Height          =   255
               Left            =   4125
               TabIndex        =   35
               Top             =   135
               Width           =   1290
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   7665
         Left            =   13395
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   300
         Width           =   12750
         _cx             =   22490
         _cy             =   13520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16185592
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin TrueOleDBGrid80.TDBGrid tdgPayrollPeriod 
            Height          =   3465
            Left            =   60
            TabIndex        =   17
            Top             =   540
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   6112
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Description"
            Columns(0).DataField=   "description"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Payroll year"
            Columns(1).DataField=   "payyear"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Month"
            Columns(2).DataField=   "paymonth"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Pay frequency"
            Columns(3).DataField=   "payfreqname"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cassification"
            Columns(4).DataField=   "classification"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "From"
            Columns(5).DataField=   "wrkdatefrom"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "To"
            Columns(6).DataField=   "wrkdateto"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   4
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   -1  'True
            Splits(0)._GSX_SAVERECORDSELECTORS=   0
            Splits(0).AllowColSelect=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            DirectionAfterTab=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Arial"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H400000&"
            _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
            _StyleDefs(17)  =   ":id=8,.fgcolor=&H0&"
            _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFF0EA&"
            _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33"
            _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HF6F8F8&"
            _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList80.TDBCombo tdbPayFreqList 
            Height          =   345
            Left            =   825
            TabIndex        =   16
            Tag             =   "Municipal"
            Top             =   150
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   609
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            _EDITHEIGHT     =   609
            _GAPHEIGHT      =   53
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Code"
            Columns(0).DataField=   "provcode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descritpion"
            Columns(1).DataField=   "provdesc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   0
            BorderStyle     =   1
            ComboStyle      =   0
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
            _PropDict       =   $"frmMDPeriod.frx":713E
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
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Named:id=33:Normal"
            _StyleDefs(41)  =   ":id=33,.parent=0"
            _StyleDefs(42)  =   "Named:id=34:Heading"
            _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(44)  =   ":id=34,.wraptext=-1"
            _StyleDefs(45)  =   "Named:id=35:Footing"
            _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   "Named:id=36:Selected"
            _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(49)  =   "Named:id=37:Caption"
            _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(51)  =   "Named:id=38:HighlightRow"
            _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=39:EvenRow"
            _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(55)  =   "Named:id=40:OddRow"
            _StyleDefs(56)  =   ":id=40,.parent=33"
            _StyleDefs(57)  =   "Named:id=41:RecordSelector"
            _StyleDefs(58)  =   ":id=41,.parent=34"
            _StyleDefs(59)  =   "Named:id=42:FilterBar"
            _StyleDefs(60)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Filter"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   660
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   7665
         Left            =   13695
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   300
         Width           =   12750
         _cx             =   22490
         _cy             =   13520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame3 
            Height          =   1350
            Left            =   180
            TabIndex        =   22
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   23
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDPeriod.frx":71E8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":7254
               Key             =   "frmMDPeriod.frx":7272
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   2
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   2
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText TDBText9 
               Height          =   300
               Left            =   1800
               TabIndex        =   24
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDPeriod.frx":72B6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":7322
               Key             =   "frmMDPeriod.frx":7340
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   2
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   2
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin TDBText6Ctl.TDBText TDBText10 
               Height          =   300
               Left            =   1800
               TabIndex        =   25
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDPeriod.frx":7384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDPeriod.frx":73F0
               Key             =   "frmMDPeriod.frx":740E
               BackColor       =   -2147483643
               EditMode        =   0
               ForeColor       =   -2147483640
               ReadOnly        =   0
               ShowContextMenu =   -1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MarginBottom    =   1
               Enabled         =   -1
               MousePointer    =   0
               Appearance      =   2
               BorderStyle     =   1
               AlignHorizontal =   0
               AlignVertical   =   2
               MultiLine       =   0
               ScrollBars      =   0
               PasswordChar    =   ""
               AllowSpace      =   -1
               Format          =   ""
               FormatMode      =   1
               AutoConvert     =   -1
               ErrorBeep       =   0
               MaxLength       =   0
               LengthAsByte    =   0
               Text            =   ""
               Furigana        =   0
               HighlightText   =   0
               IMEMode         =   0
               IMEStatus       =   0
               DropWndWidth    =   0
               DropWndHeight   =   0
               ScrollBarMode   =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mode of Payment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   30
               TabIndex        =   28
               Top             =   615
               Width           =   1635
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   345
               TabIndex        =   27
               Top             =   300
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   660
               TabIndex        =   26
               Top             =   945
               Width           =   990
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   29
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "frmMDPeriod.frx":7452
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDPeriod.frx":74BE
            Key             =   "frmMDPeriod.frx":74DC
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   -1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   2
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   2
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   3600
            Left            =   180
            TabIndex        =   30
            Top             =   2115
            Width           =   7275
            _cx             =   12832
            _cy             =   6350
            Appearance      =   3
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   13431287
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmMDPeriod.frx":7520
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
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   915
            TabIndex        =   31
            Top             =   240
            Width           =   915
         End
      End
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   44
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Edit"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   0
         Left            =   75
         TabIndex        =   45
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&New"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   2
         Left            =   2385
         TabIndex        =   46
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Delete"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   3
         Left            =   3540
         TabIndex        =   47
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   4
         Left            =   4695
         TabIndex        =   48
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Print"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   5
         Left            =   5850
         TabIndex        =   49
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   16185592
      End
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   7875
      TabIndex        =   50
      Top             =   60
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   609
      Caption         =   "Title"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmMDPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPayrollPeriod         As ADODB.Recordset
Dim rsLoanDed               As ADODB.Recordset

Dim mSort                   As String

Private Sub cmdMenu_Click(Index As Integer)
  Select Case Index
    Case 0: AddSave_Button_Clicked
    Case 1: EditUpdate_Button_Clicked
    Case 2:
    Case 3: Cancel_Clicked
    Case 4:
    Case 5: Unload Me
  End Select
End Sub

Private Sub Form_Load()

    Dim rsTmp         As ADODB.Recordset
    Dim rsTmpPayFreq  As ADODB.Recordset
    Dim i             As Integer
    
    Add_MDIButton Me.Name, TitleBar.Caption
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 12
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "January"
          Case 2: .Fields("description") = "February"
          Case 3: .Fields("description") = "March"
          Case 4: .Fields("description") = "April"
          Case 5: .Fields("description") = "May"
          Case 6: .Fields("description") = "June"
          Case 7: .Fields("description") = "July"
          Case 8: .Fields("description") = "August"
          Case 9: .Fields("description") = "September"
          Case 10: .Fields("description") = "October"
          Case 11: .Fields("description") = "November"
          Case 12: .Fields("description") = "December"
        End Select
        .Update
      Next
    End With
    
    With tdbMonth
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 5
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "First"
          Case 2: .Fields("description") = "Second"
          Case 3: .Fields("description") = "Third"
          Case 4: .Fields("description") = "Fourth"
          Case 5: .Fields("description") = "Fifth"
        End Select
        .Update
      Next
    End With
    
    With tdbClassification
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    NetOpen rsTmpPayFreq, "select * from payfrequency order by payfreqname"
    With rsTmpPayFreq
      CreateTmpDB rsTmp
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          rsTmp.AddNew
          rsTmp.Fields("code") = !payfreqcode
          rsTmp.Fields("Description") = !payfreqname
          rsTmp.Update
          .MoveNext
        Loop
      End If
      rsTmp.AddNew
      rsTmp.Fields("code") = "All"
      rsTmp.Fields("description") = "All"
    End With
    
    With tdbPayFreqList
      .BoundColumn = "code"
      .ListField = "description"
      .Columns(0).DataField = "code"
      .Columns(1).DataField = "description"
      .RowSource = rsTmp
    End With
      
    Set rsTmp = Nothing
    
    bind_tdb ConMain, tdbPayFrequency, "select payfreqcode,payfreqname from payfrequency order by payfreqname", "payfreqname", "payfreqcode"

    
    NetOpen rsPayrollPeriod, "select x1.*,x2.payfreqname from payrollperiod x1 " & _
                                 "left outer join payfrequency x2 on x1.payfreqcode = x2.payfreqcode "
    If rsPayrollPeriod.RecordCount > 0 Then
      rsPayrollPeriod.MoveLast
    End If
    Set tdgPayrollPeriod.DataSource = rsPayrollPeriod
    mSort = "percode"
    
    tdbPayFreqList.BoundText = "All"
    tabPayrollPeriod.CurrTab = 0
    cmdMenu_Click 3

End Sub

Private Sub AddSave_Button_Clicked()

  If cmdMenu(0).Caption = "&New" Then
    
    tabPayrollPeriod.CurrTab = 0
    Lock_Button "TFFTFF", cmdMenu, 5
    cmdMenu(0).Caption = "&Save"
    ClearText
    Lock_Tab "TF", tabPayrollPeriod, 1
    FraInfo.Enabled = True
    txtPayyear.SetFocus
    
  Else
  
    If Not IsDate(txtPayyear.Text) Then
      MsgBox "Invalid year format.", vbExclamation + vbOKOnly
      txtPayyear.SetFocus
      Exit Sub
    End If
     
    If Trim(tdbMonth.Text) = "" Or IsNull(tdbMonth.SelectedItem) Or tdbMonth.ApproxCount = 0 Then
      MsgBox "Please select a month.", vbExclamation + vbOKOnly
      tdbMonth.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbPayFrequency.Text) = "" Or IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
      MsgBox "Please select a pay frequency.", vbExclamation + vbOKOnly
      tdbPayFrequency.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbClassification.Text) = "" Or IsNull(tdbClassification.SelectedItem) Or tdbClassification.ApproxCount = 0 Then
      MsgBox "Please select a classification.", vbExclamation + vbOKOnly
      tdbClassification.SetFocus
      Exit Sub
    End If
    
    If Trim(txtDescription.Text) = "" Then
      MsgBox "Please provide a description.", vbExclamation + vbOKOnly
      txtDescription.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtFrom.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtFrom.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtTo.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtTo.SetFocus
      Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    txtPayPeriod.Text = LastCode("PayrollPeriod")
    ConMain.Execute "insert into payrollperiod(payyear,paymonth,percode,description,payfreqcode," & _
                    "classification,wrkdatefrom,wrkdateto,sssdaily,sssmonthly," & _
                    "phdaily,phmonthly,taxdaily,taxmonthly,hdmfdaily," & _
                    "hdmfmonthly,lastotcode) values " & _
                    "('" & txtPayyear.Text & "','" & tdbMonth.Text & "'," & txtPayPeriod.Text & ",'" & txtDescription.Text & "','" & tdbPayFrequency.BoundText & "', " & _
                    "'" & tdbClassification.Text & "','" & Format(txtFrom.Text, "YYYY-MM-DD") & "','" & Format(txtTo.Text, "YYYY-MM-DD") & "','" & IIf(chkSSSDaily.Value = 0, "N", "Y") & "','" & IIf(chkSSSMonthly.Value = 0, "N", "Y") & "'," & _
                    "'" & IIf(chkPHDaily.Value = 0, "N", "Y") & "','" & IIf(chkPHMonthly.Value = 0, "N", "Y") & "','" & IIf(chkTaxDaily.Value = 0, "N", "Y") & "','" & IIf(chkTaxMonthly.Value = 0, "N", "Y") & "','" & IIf(chkHDMFDaily.Value = 0, "N", "Y") & "','" & _
                    IIf(chkHDMFMonthly.Value = 0, "N", "Y") & "',1) "

    ConMain.Execute "delete from payrollperiodloandedallow where percode = " & rsPayrollPeriod!percode & ""
    With rsLoanDed
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !allow <> 0 Then
                    ConMain.Execute "insert into payrollperiodloandedallow (percode,loantypescode,allow) values (" & _
                             rsPayrollPeriod!percode & "," & !loantypescode & "," & IIf(!allow <> 0, 1, 0) & ")"
                End If
                .MoveNext
            Loop
        End If
    End With
    
    ConMain.CommitTrans
    tdbPayFreqList.BoundText = "All"
    rsPayrollPeriod.Requery
    rsPayrollPeriod.Find "percode = '" & txtPayPeriod.Text & "'"
        
    cmdMenu_Click 3
      
  End If
  
End Sub

Private Sub EditUpdate_Button_Clicked()

  If cmdMenu(1).Caption = "&Edit" Then
    
    tabPayrollPeriod.CurrTab = 0
    Lock_Button "FTFTFF", cmdMenu, 5
    cmdMenu(1).Caption = "&Update"
    Lock_Tab "TF", tabPayrollPeriod, 1
    FraInfo.Enabled = True
    txtPayyear.SetFocus
  Else
  
  
    If Trim(txtPayyear.Text) = "" Then
      MsgBox "Payroll year is blank.", vbExclamation + vbOKOnly
      txtPayyear.SetFocus
      Exit Sub
    End If
     
    If Trim(tdbMonth.Text) = "" Or IsNull(tdbMonth.SelectedItem) Or tdbMonth.ApproxCount = 0 Then
      MsgBox "Please select a month.", vbExclamation + vbOKOnly
      tdbMonth.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbPayFrequency.Text) = "" Or IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
      MsgBox "Please select a pay frequency.", vbExclamation + vbOKOnly
      tdbPayFrequency.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbClassification.Text) = "" Or IsNull(tdbClassification.SelectedItem) Or tdbClassification.ApproxCount = 0 Then
      MsgBox "Please select a classification.", vbExclamation + vbOKOnly
      tdbClassification.SetFocus
      Exit Sub
    End If
    
    If Trim(txtDescription.Text) = "" Then
      MsgBox "Please provide description.", vbExclamation + vbOKOnly
      txtDescription.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtFrom.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtFrom.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtTo.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtTo.SetFocus
      Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    ConMain.Execute "update payrollperiod set payyear = '" & txtPayyear.Text & "', paymonth = '" & tdbMonth.Text & "', payfreqcode = '" & tdbPayFrequency.BoundText & "', description = '" & txtDescription.Text & "'," & _
                       "classification = '" & tdbClassification.Text & "',  wrkdatefrom = '" & Format(txtFrom.Text, "YYYY-MM-DD") & "', wrkdateto = '" & Format(txtTo.Text, "YYYY-MM-DD") & "', " & _
                       "sssdaily = '" & IIf(chkSSSDaily.Value = 0, "N", "Y") & "', sssmonthly = '" & IIf(chkSSSMonthly.Value = 0, "N", "Y") & "', " & _
                       "phdaily = '" & IIf(chkPHDaily.Value = 0, "N", "Y") & "', phmonthly = '" & IIf(chkPHMonthly.Value = 0, "N", "Y") & "', " & _
                       "taxdaily = '" & IIf(chkTaxDaily.Value = 0, "N", "Y") & "', taxmonthly = '" & IIf(chkTaxMonthly.Value = 0, "N", "Y") & "', " & _
                       "hdmfdaily = '" & IIf(chkHDMFDaily.Value = 0, "N", "Y") & "', hdmfmonthly = '" & IIf(chkHDMFMonthly.Value = 0, "N", "Y") & "' where percode = " & txtPayPeriod.Text & ""

    ConMain.Execute "delete from payrollperiodloandedallow where percode = " & rsPayrollPeriod!percode & ""
    With rsLoanDed
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !allow <> 0 Then
                    ConMain.Execute "insert into payrollperiodloandedallow (percode,loantypescode,allow) values (" & _
                             rsPayrollPeriod!percode & "," & !loantypescode & "," & IIf(!allow <> 0, 1, 0) & ")"
                End If
                .MoveNext
            Loop
        End If
    End With
    
    ConMain.CommitTrans
    tdbPayFreqList.BoundText = "All"
    rsPayrollPeriod.Requery
    rsPayrollPeriod.Find "percode = '" & txtPayPeriod.Text & "'"
        
    cmdMenu_Click 3
    
  End If
  
End Sub

Private Sub Cancel_Clicked()

  If rsPayrollPeriod.RecordCount > 0 Then
    Lock_Button "TTTFTT", cmdMenu, 5
  Else
    Lock_Button "TFFFTT", cmdMenu, 5
  End If

  cmdMenu(0).Caption = "&New"
  cmdMenu(1).Caption = "&Edit"
  
  tabPayrollPeriod.Enabled = True
  FraInfo.Enabled = False
  Lock_Tab "TT", tabPayrollPeriod, 1
  tdgPayrollPeriod_RowColChange 0, 0
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
      .Top = TitleBar.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With tabPayrollPeriod
      .Top = frabutton.Height + frabutton.Top
      .Left = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight - .Top
    End With
    
    With tdgPayrollPeriod
      .Left = 150
      .Width = Me.ScaleWidth - 300
      .Height = tabPayrollPeriod.Height - (.Top + 800)
    End With
    
End Sub

Private Sub tabPayrollPeriod_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If NewTab = 0 Then
        Create_TmpLoanDed
    End If
End Sub

Private Sub tdbClassification_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbClassification, tdbClassification.RowSource, tdbClassification.Text
  End If
End Sub

Private Sub tdbMonth_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbMonth, tdbMonth.RowSource, tdbMonth.Text
  End If
End Sub

Private Sub tdbPayFreqList_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbPayFreqList, tdbPayFreqList.RowSource, tdbPayFreqList.Text
  End If
End Sub

Private Sub tdbPayFreqList_ItemChange()
  
  rsPayrollPeriod.Filter = ""
  If tdbPayFreqList.BoundText <> "All" Then
    rsPayrollPeriod.Filter = "payfreqcode = '" & tdbPayFreqList.BoundText & "'"
  End If
  If Not rsPayrollPeriod.EOF Then
    rsPayrollPeriod.Sort = "percode"
    rsPayrollPeriod.MoveLast
  End If
  
End Sub

Private Sub tdbPayFrequency_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbPayFrequency, tdbPayFrequency.RowSource, tdbPayFrequency.Text
    tdbPayFreqList_ItemChange
  End If
End Sub

Private Sub tdgPayrollPeriod_HeadClick(ByVal ColIndex As Integer)
  
  If ColIndex <= 2 Then
    If rsPayrollPeriod.RecordCount > 0 Then
      mSort = tdgPayrollPeriod.Columns(ColIndex).DataField
      rsPayrollPeriod.Sort = mSort
    End If
  End If
  
End Sub

Private Sub tdgPayrollPeriod_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  
  With rsPayrollPeriod

    If .RecordCount > 0 Then
    
      If LastRow <> tdgPayrollPeriod.Row + 1 Then
        txtPayyear.Text = !payyear
        tdbMonth.Text = !paymonth
        tdbPayFrequency.BoundText = !payfreqcode
        tdbClassification.Text = !classification
        txtPayPeriod.Text = !percode
        txtDescription.Text = !Description
        txtFrom.Text = Format(!wrkdatefrom, "MM/DD/YYYY")
        txtTo.Text = Format(!wrkdateto, "MM/DD/YYYY")
        chkSSSDaily.Value = IIf(!sssdaily = "N", 0, 1)
        chkSSSMonthly.Value = IIf(!sssmonthly = "N", 0, 1)
        chkPHDaily.Value = IIf(!phdaily = "N", 0, 1)
        chkPHMonthly.Value = IIf(!phmonthly = "N", 0, 1)
        chkTaxDaily.Value = IIf(!taxdaily = "N", 0, 1)
        chkTaxMonthly.Value = IIf(!taxmonthly = "N", 0, 1)
        chkHDMFDaily.Value = IIf(!hdmfdaily = "N", 0, 1)
        chkHDMFMonthly.Value = IIf(!hdmfmonthly = "N", 0, 1)
      End If

    End If
    
  End With
  
End Sub

Private Sub ClearText()

    txtPayyear.Text = ""
    tdbMonth.BoundText = ""
    tdbPayFrequency.BoundText = ""
    tdbClassification.BoundText = ""
    txtPayPeriod.Text = ""
    txtDescription.Text = ""
    txtFrom.Text = Format(Now, "MM/DD/YYYY")
    txtTo.Text = Format(Now, "MM/DD/YYYY")
    chkSSSDaily.Value = 0
    chkSSSMonthly.Value = 0
    chkPHDaily.Value = 0
    chkPHMonthly.Value = 0
    chkTaxDaily.Value = 0
    chkTaxMonthly.Value = 0
    chkHDMFDaily.Value = 0
    chkHDMFMonthly.Value = 0
    
End Sub

Public Sub Create_TmpLoanDed()

    Dim rsLoanDedTmp        As ADODB.Recordset

    Set rsLoanDed = Nothing
    Set rsLoanDed = New ADODB.Recordset
    
    With rsLoanDed
      .Fields.Append "loantypescode", adVarChar, 7
      .Fields.Append "loantypesname", adVarChar, 50
      .Fields.Append "allow", adInteger
      .Open
    End With
    
    Set tdgLeaveLimit.DataSource = rsLoanDed
    
    NetOpen rsLoanDedTmp, "select x1.loantypescode,x1.loantypesname," & _
                        "(select allow from payrollperiodloandedallow where loantypescode = x1.loantypescode and percode = " & rsPayrollPeriod!percode & " ) allow " & _
                        "from loantypes x1 order by x1.loantypesname "

    With rsLoanDedTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                rsLoanDed.AddNew
                rsLoanDed.Fields("loantypescode") = !loantypescode
                rsLoanDed.Fields("loantypesname") = !loantypesname
                rsLoanDed.Fields("allow") = IIf(IsNull(!allow), 0, !allow)
                rsLoanDed.Update
                .MoveNext
            Loop
        End If
    End With
    
End Sub

