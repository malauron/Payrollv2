VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDEmployee 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   13065
   Tag             =   "Master data - Employee"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraHolder 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4830
      Left            =   90
      TabIndex        =   73
      Top             =   870
      Width           =   12945
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   30
         TabIndex        =   151
         Top             =   -60
         Width           =   12915
      End
      Begin VB.Frame fraSearchEmployee 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   30
         TabIndex        =   148
         Top             =   0
         Width           =   12915
         Begin lvButton.lvButtons_H cmdShowRegistration 
            Height          =   300
            Left            =   1800
            TabIndex        =   81
            Top             =   135
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   529
            Caption         =   "..."
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   33023
            cFHover         =   33023
            cBhover         =   16777215
            cGradient       =   16777215
            Gradient        =   4
            CapStyle        =   2
            Mode            =   0
            Value           =   0   'False
            ImgAlign        =   1
            cBack           =   14737632
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH EMPLOYEE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   -60
            TabIndex        =   150
            Top             =   210
            Width           =   1815
         End
         Begin VB.Label lblEmployeeName 
            Alignment       =   2  'Center
            BackColor       =   &H00F6F8F8&
            BackStyle       =   0  'Transparent
            Caption         =   "0000059  - MARK ANTHONY M. LAURON"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   375
            Left            =   2235
            TabIndex        =   149
            Top             =   90
            Width           =   10515
         End
      End
      Begin C1SizerLibCtl.C1Tab tabEmployee 
         Height          =   3675
         Left            =   30
         TabIndex        =   74
         Top             =   495
         Width           =   12900
         _cx             =   22754
         _cy             =   6482
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         BackColor       =   16777215
         ForeColor       =   4210752
         FrontTabColor   =   14737632
         BackTabColor    =   14737632
         TabOutlineColor =   12632256
         FrontTabForeColor=   4210752
         Caption         =   "Employment Information|Personal Information|Documents|Salary"
         Align           =   0
         CurrTab         =   3
         FirstTab        =   0
         Style           =   4
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   4
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3360
            Left            =   15
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   300
            Width           =   12870
            _cx             =   22701
            _cy             =   5927
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14737632
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
            Begin VB.Frame frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   3405
               Index           =   1
               Left            =   45
               TabIndex        =   76
               Top             =   -60
               Width           =   12810
               Begin VB.Frame frame1 
                  BackColor       =   &H00E0E0E0&
                  Height          =   3405
                  Index           =   2
                  Left            =   7980
                  TabIndex        =   77
                  Top             =   0
                  Width           =   60
               End
               Begin lvButton.lvButtons_H cmdNonTaxAllow 
                  Height          =   420
                  Left            =   855
                  TabIndex        =   78
                  Top             =   885
                  Width           =   5880
                  _ExtentX        =   10372
                  _ExtentY        =   741
                  Caption         =   "Non-Taxable Allowances"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin TDBNumber6Ctl.TDBNumber txtMonthly_Rate 
                  Height          =   300
                  Left            =   8520
                  TabIndex        =   55
                  Top             =   1755
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":0000
                  Caption         =   "frmMDEmployee.frx":0020
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":0086
                  Keys            =   "frmMDEmployee.frx":00A4
                  Spin            =   "frmMDEmployee.frx":00EE
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,###,##0.0000"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,###,##0.0000"
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
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TrueOleDBList80.TDBCombo tdbRateType 
                  Height          =   300
                  Left            =   9960
                  TabIndex        =   54
                  Top             =   630
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":0116
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=0,.fontsize=825"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin TDBNumber6Ctl.TDBNumber txtMealAllow 
                  Height          =   300
                  Left            =   10605
                  TabIndex        =   58
                  Top             =   1755
                  Width           =   1530
                  _Version        =   65536
                  _ExtentX        =   2699
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":01C0
                  Caption         =   "frmMDEmployee.frx":01E0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":0246
                  Keys            =   "frmMDEmployee.frx":0264
                  Spin            =   "frmMDEmployee.frx":02AE
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,###,###,##0.00"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,###,###,##0.00"
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
                  ValueVT         =   1999437829
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber txtFixedEarnings 
                  Height          =   300
                  Left            =   10605
                  TabIndex        =   59
                  Top             =   2295
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":02D6
                  Caption         =   "frmMDEmployee.frx":02F6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":035C
                  Keys            =   "frmMDEmployee.frx":037A
                  Spin            =   "frmMDEmployee.frx":03C4
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,###,###,##0.00"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,###,###,##0.00"
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
                  ValueVT         =   1999437829
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber txtDaily_Rate 
                  Height          =   300
                  Left            =   8550
                  TabIndex        =   56
                  Top             =   2295
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":03EC
                  Caption         =   "frmMDEmployee.frx":040C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":0472
                  Keys            =   "frmMDEmployee.frx":0490
                  Spin            =   "frmMDEmployee.frx":04DA
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,##0.0000000"
                  EditMode        =   0
                  Enabled         =   0
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,##0.0000000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   99999
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   0
                  Separator       =   ","
                  ShowContextMenu =   -1
                  ValueVT         =   1999437829
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber txtHourly_Rate 
                  Height          =   300
                  Left            =   8550
                  TabIndex        =   57
                  Top             =   2835
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":0502
                  Caption         =   "frmMDEmployee.frx":0522
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":0588
                  Keys            =   "frmMDEmployee.frx":05A6
                  Spin            =   "frmMDEmployee.frx":05F0
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,##0.0000000"
                  EditMode        =   0
                  Enabled         =   0
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,##0.0000000"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   99999
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   0
                  Separator       =   ","
                  ShowContextMenu =   -1
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TrueOleDBList80.TDBCombo tdbPayFrequency 
                  Height          =   300
                  Left            =   9960
                  TabIndex        =   53
                  Top             =   255
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":0618
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=0,.fontsize=825"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin lvButton.lvButtons_H lvButtons_H10 
                  Height          =   420
                  Left            =   855
                  TabIndex        =   79
                  Top             =   1380
                  Visible         =   0   'False
                  Width           =   5880
                  _ExtentX        =   10372
                  _ExtentY        =   741
                  Caption         =   "Loans Monitoring"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin lvButton.lvButtons_H lvButtons_H11 
                  Height          =   420
                  Left            =   855
                  TabIndex        =   80
                  Top             =   1875
                  Visible         =   0   'False
                  Width           =   5880
                  _ExtentX        =   10372
                  _ExtentY        =   741
                  Caption         =   "Leaves Monitoring"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin TDBNumber6Ctl.TDBNumber txtCola 
                  Height          =   300
                  Left            =   10620
                  TabIndex        =   166
                  Top             =   2835
                  Width           =   1515
                  _Version        =   65536
                  _ExtentX        =   2672
                  _ExtentY        =   529
                  Calculator      =   "frmMDEmployee.frx":06C2
                  Caption         =   "frmMDEmployee.frx":06E2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":0748
                  Keys            =   "frmMDEmployee.frx":0766
                  Spin            =   "frmMDEmployee.frx":07B0
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "###,###,###,##0.00"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   4210752
                  Format          =   "###,###,###,##0.00"
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
                  ValueVT         =   1999437829
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TrueOleDBList80.TDBCombo tdbWrkDays 
                  Height          =   300
                  Left            =   10920
                  TabIndex        =   168
                  Top             =   1020
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":07D8
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=0,.fontsize=825"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Working Days/Month"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   32
                  Left            =   8520
                  TabIndex        =   169
                  Top             =   1080
                  Width           =   2985
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "C.O.L.A."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   1
                  Left            =   10620
                  TabIndex        =   167
                  Top             =   2625
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Rate Type"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   10
                  Left            =   8520
                  TabIndex        =   88
                  Top             =   690
                  Width           =   1725
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pay Frequency"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   2
                  Left            =   8535
                  TabIndex        =   87
                  Top             =   315
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Daily Rate"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   9
                  Left            =   8550
                  TabIndex        =   86
                  Top             =   2085
                  Width           =   1725
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Monthly Rate"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   3
                  Left            =   8520
                  TabIndex        =   85
                  Top             =   1545
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hourly Rate"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   11
                  Left            =   8565
                  TabIndex        =   84
                  Top             =   2625
                  Width           =   1725
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Meal Allowance/ Day"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   4
                  Left            =   10620
                  TabIndex        =   83
                  Top             =   1545
                  Width           =   1905
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fixed Earnings"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   5
                  Left            =   10605
                  TabIndex        =   82
                  Top             =   2085
                  Width           =   1725
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3360
            Left            =   -13485
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   300
            Width           =   12870
            _cx             =   22701
            _cy             =   5927
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14737632
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
               BackColor       =   &H00E0E0E0&
               Height          =   3390
               Left            =   30
               TabIndex        =   90
               Top             =   -60
               Width           =   12810
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   690
                  Left            =   2955
                  TabIndex        =   91
                  Top             =   150
                  Width           =   9810
                  Begin VB.OptionButton optTaxFixed 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Fixed"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   7605
                     TabIndex        =   32
                     Top             =   315
                     Width           =   750
                  End
                  Begin VB.OptionButton optTaxAuto 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Auto Deduct"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   3015
                     TabIndex        =   30
                     Top             =   315
                     Width           =   1395
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtTaxAmt 
                     Height          =   300
                     Left            =   8430
                     TabIndex        =   33
                     Top             =   330
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":0882
                     Caption         =   "frmMDEmployee.frx":08A2
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0908
                     Keys            =   "frmMDEmployee.frx":0926
                     Spin            =   "frmMDEmployee.frx":0970
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBText6Ctl.TDBText txtTinno 
                     Height          =   300
                     Left            =   795
                     TabIndex        =   29
                     Top             =   300
                     Width           =   2160
                     _Version        =   65536
                     _ExtentX        =   3810
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":0998
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0A04
                     Key             =   "frmMDEmployee.frx":0A22
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin TDBText6Ctl.TDBText txtWT 
                     Height          =   300
                     Left            =   4455
                     TabIndex        =   31
                     Top             =   315
                     Width           =   2700
                     _Version        =   65536
                     _ExtentX        =   4762
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":0A66
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0AD2
                     Key             =   "frmMDEmployee.frx":0AF0
                     BackColor       =   16777215
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin lvButton.lvButtons_H cmdWT 
                     Height          =   315
                     Left            =   7170
                     TabIndex        =   69
                     Top             =   315
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   556
                     Caption         =   "..."
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
                     cBack           =   14737632
                  End
                  Begin TrueOleDBList80.TDBCombo tdbTmp2 
                     Bindings        =   "frmMDEmployee.frx":0B34
                     DataMember      =   "tdbJob"
                     Height          =   300
                     Left            =   0
                     TabIndex        =   154
                     Top             =   615
                     Visible         =   0   'False
                     Width           =   2835
                     _ExtentX        =   5001
                     _ExtentY        =   529
                     _LayoutType     =   4
                     _RowHeight      =   -2147483647
                     _WasPersistedAsPixels=   0
                     _DropdownWidth  =   0
                     _EDITHEIGHT     =   529
                     _GAPHEIGHT      =   53
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns(1)._VlistStyle=   0
                     Columns(1)._MaxComboItems=   5
                     Columns(1).DataField=   ""
                     Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   2
                     Splits(0)._UserFlags=   0
                     Splits(0).ExtendRightColumn=   -1  'True
                     Splits(0).AllowRowSizing=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                     Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                     Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                     Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                     Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                     Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                     Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                     Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                     Splits.Count    =   1
                     Appearance      =   0
                     BorderStyle     =   1
                     ComboStyle      =   0
                     AutoCompletion  =   0   'False
                     LimitToList     =   0   'False
                     ColumnHeaders   =   -1  'True
                     ColumnFooters   =   0   'False
                     DataMode        =   0
                     DefColWidth     =   0
                     Enabled         =   -1  'True
                     HeadLines       =   0
                     FootLines       =   1
                     RowDividerStyle =   0
                     Caption         =   ""
                     EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                     LayoutName      =   ""
                     LayoutFileName  =   ""
                     MultipleLines   =   0
                     EmptyRows       =   -1  'True
                     CellTips        =   0
                     AutoSize        =   0   'False
                     ListField       =   ""
                     BoundColumn     =   ""
                     IntegralHeight  =   0   'False
                     CellTipsWidth   =   0
                     CellTipsDelay   =   1000
                     AutoDropdown    =   0   'False
                     RowTracking     =   -1  'True
                     RightToLeft     =   0   'False
                     RowMember       =   ""
                     MouseIcon       =   0
                     MouseIcon.vt    =   3
                     MousePointer    =   0
                     MatchEntryTimeout=   2000
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     AnimateWindow   =   0
                     AnimateWindowDirection=   0
                     AnimateWindowTime=   200
                     AnimateWindowClose=   0
                     DropdownPosition=   0
                     Locked          =   0   'False
                     ScrollTrack     =   0   'False
                     RowDividerColor =   14215660
                     RowSubDividerColor=   14215660
                     AddItemSeparator=   ";"
                     _PropDict       =   $"frmMDEmployee.frx":0B45
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                     _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
                  Begin VB.Label Label4 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "TAX"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Index           =   0
                     Left            =   45
                     TabIndex        =   158
                     Top             =   60
                     Width           =   330
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "T.I.N. No"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     Left            =   90
                     TabIndex        =   92
                     Top             =   360
                     Width           =   1095
                  End
               End
               Begin lvButton.lvButtons_H lvButtons_H13 
                  Height          =   585
                  Index           =   0
                  Left            =   90
                  TabIndex        =   93
                  Top             =   240
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   1032
                  Caption         =   "Witholding Tax"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin lvButton.lvButtons_H lvButtons_H13 
                  Height          =   585
                  Index           =   1
                  Left            =   90
                  TabIndex        =   99
                  Top             =   840
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   1032
                  Caption         =   "SSS"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin lvButton.lvButtons_H lvButtons_H13 
                  Height          =   600
                  Index           =   2
                  Left            =   90
                  TabIndex        =   108
                  Top             =   1440
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   1058
                  Caption         =   "HDMF"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin lvButton.lvButtons_H lvButtons_H13 
                  Height          =   585
                  Index           =   3
                  Left            =   90
                  TabIndex        =   109
                  Top             =   2055
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   1032
                  Caption         =   "PhilHealth"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin lvButton.lvButtons_H cmdAddBankAccount 
                  Height          =   585
                  Left            =   90
                  TabIndex        =   113
                  Top             =   2655
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   1032
                  Caption         =   "Set Bank Account"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   4194304
                  LockHover       =   2
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   690
                  Left            =   2955
                  TabIndex        =   94
                  Top             =   750
                  Width           =   9810
                  Begin VB.OptionButton optSSSFixed 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Fixed"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   4470
                     TabIndex        =   36
                     Top             =   300
                     Width           =   750
                  End
                  Begin VB.OptionButton optSSSAuto 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Auto Deduct"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   3015
                     TabIndex        =   35
                     Top             =   300
                     Width           =   1395
                  End
                  Begin TDBText6Ctl.TDBText txtSSSno 
                     Height          =   300
                     Left            =   780
                     TabIndex        =   34
                     Top             =   300
                     Width           =   2160
                     _Version        =   65536
                     _ExtentX        =   3810
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":0BEF
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0C5B
                     Key             =   "frmMDEmployee.frx":0C79
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin TDBNumber6Ctl.TDBNumber txtSSSAmt 
                     Height          =   300
                     Left            =   5340
                     TabIndex        =   37
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":0CBD
                     Caption         =   "frmMDEmployee.frx":0CDD
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0D43
                     Keys            =   "frmMDEmployee.frx":0D61
                     Spin            =   "frmMDEmployee.frx":0DAB
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtSssEr 
                     Height          =   300
                     Left            =   6705
                     TabIndex        =   38
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":0DD3
                     Caption         =   "frmMDEmployee.frx":0DF3
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0E59
                     Keys            =   "frmMDEmployee.frx":0E77
                     Spin            =   "frmMDEmployee.frx":0EC1
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtSssEc 
                     Height          =   300
                     Left            =   8055
                     TabIndex        =   39
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":0EE9
                     Caption         =   "frmMDEmployee.frx":0F09
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":0F6F
                     Keys            =   "frmMDEmployee.frx":0F8D
                     Spin            =   "frmMDEmployee.frx":0FD7
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "SSS"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Index           =   1
                     Left            =   45
                     TabIndex        =   159
                     Top             =   75
                     Width           =   330
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acct No"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     Left            =   75
                     TabIndex        =   98
                     Top             =   345
                     Width           =   1260
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "SSS EE:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   25
                     Left            =   5355
                     TabIndex        =   97
                     Top             =   120
                     Width           =   1260
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "SSS ER:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   26
                     Left            =   6720
                     TabIndex        =   96
                     Top             =   120
                     Width           =   1260
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "SSS EC:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   27
                     Left            =   8085
                     TabIndex        =   95
                     Top             =   120
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   705
                  Left            =   2955
                  TabIndex        =   100
                  Top             =   1350
                  Width           =   9810
                  Begin VB.OptionButton optHDMFFixed 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Fixed"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   4515
                     TabIndex        =   42
                     Top             =   330
                     Width           =   750
                  End
                  Begin VB.OptionButton optHDMFAuto 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Auto Deduct"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   3030
                     TabIndex        =   41
                     Top             =   330
                     Width           =   1395
                  End
                  Begin TDBText6Ctl.TDBText txtHDMFNo 
                     Height          =   300
                     Left            =   780
                     TabIndex        =   40
                     Top             =   315
                     Width           =   2160
                     _Version        =   65536
                     _ExtentX        =   3810
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":0FFF
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":106B
                     Key             =   "frmMDEmployee.frx":1089
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin TDBNumber6Ctl.TDBNumber txtHdmfAmt 
                     Height          =   300
                     Left            =   5355
                     TabIndex        =   43
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":10CD
                     Caption         =   "frmMDEmployee.frx":10ED
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":1153
                     Keys            =   "frmMDEmployee.frx":1171
                     Spin            =   "frmMDEmployee.frx":11BB
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtHDMFEr 
                     Height          =   300
                     Left            =   6720
                     TabIndex        =   44
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":11E3
                     Caption         =   "frmMDEmployee.frx":1203
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":1269
                     Keys            =   "frmMDEmployee.frx":1287
                     Spin            =   "frmMDEmployee.frx":12D1
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "HDMF"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Index           =   2
                     Left            =   45
                     TabIndex        =   160
                     Top             =   60
                     Width           =   480
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acct No"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     Left            =   60
                     TabIndex        =   103
                     Top             =   360
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "HDMF EE"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   28
                     Left            =   5370
                     TabIndex        =   102
                     Top             =   120
                     Width           =   1260
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "HDMF ER"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   29
                     Left            =   6720
                     TabIndex        =   101
                     Top             =   120
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   690
                  Left            =   2955
                  TabIndex        =   104
                  Top             =   1965
                  Width           =   9810
                  Begin VB.OptionButton optPhilHAuto 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Auto Deduct"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   3045
                     TabIndex        =   46
                     Top             =   300
                     Width           =   1395
                  End
                  Begin VB.OptionButton optPhilHFixed 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Fixed"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     Left            =   4530
                     TabIndex        =   47
                     Top             =   300
                     Width           =   750
                  End
                  Begin TDBText6Ctl.TDBText txtPhilHNo 
                     Height          =   300
                     Left            =   780
                     TabIndex        =   45
                     Top             =   315
                     Width           =   2160
                     _Version        =   65536
                     _ExtentX        =   3810
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":12F9
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":1365
                     Key             =   "frmMDEmployee.frx":1383
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin TDBNumber6Ctl.TDBNumber txtPhilHAmt 
                     Height          =   300
                     Left            =   5370
                     TabIndex        =   48
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":13C7
                     Caption         =   "frmMDEmployee.frx":13E7
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":144D
                     Keys            =   "frmMDEmployee.frx":146B
                     Spin            =   "frmMDEmployee.frx":14B5
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtPhilEr 
                     Height          =   300
                     Left            =   6735
                     TabIndex        =   49
                     Top             =   315
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployee.frx":14DD
                     Caption         =   "frmMDEmployee.frx":14FD
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":1563
                     Keys            =   "frmMDEmployee.frx":1581
                     Spin            =   "frmMDEmployee.frx":15CB
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "#,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   4210752
                     Format          =   "#,##0.00"
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
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "PHILHEALTH"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Index           =   3
                     Left            =   45
                     TabIndex        =   161
                     Top             =   75
                     Width           =   1050
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acct No"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     Left            =   60
                     TabIndex        =   107
                     Top             =   375
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "PhilH EE"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   30
                     Left            =   5385
                     TabIndex        =   106
                     Top             =   120
                     Width           =   1260
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "PhilH ER"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Index           =   31
                     Left            =   6750
                     TabIndex        =   105
                     Top             =   120
                     Width           =   1260
                  End
               End
               Begin VB.Frame Frame9 
                  BackColor       =   &H00E0E0E0&
                  Height          =   690
                  Left            =   2955
                  TabIndex        =   110
                  Top             =   2565
                  Width           =   9810
                  Begin VB.CheckBox chkSalToBank 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "for PAYROLL CREDIT UPLOAD"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   315
                     Left            =   7020
                     TabIndex        =   52
                     Top             =   315
                     Width           =   2715
                  End
                  Begin TDBText6Ctl.TDBText txtBankAcctNo 
                     Height          =   300
                     Left            =   4905
                     TabIndex        =   51
                     Top             =   315
                     Width           =   2010
                     _Version        =   65536
                     _ExtentX        =   3545
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":15F3
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":165F
                     Key             =   "frmMDEmployee.frx":167D
                     BackColor       =   -2147483643
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin TDBText6Ctl.TDBText txtBank 
                     Height          =   300
                     Left            =   780
                     TabIndex        =   50
                     Top             =   315
                     Width           =   2880
                     _Version        =   65536
                     _ExtentX        =   5080
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployee.frx":16C1
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployee.frx":172D
                     Key             =   "frmMDEmployee.frx":174B
                     BackColor       =   16777215
                     EditMode        =   0
                     ForeColor       =   4210752
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
                     MaxLength       =   20
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
                  Begin lvButton.lvButtons_H cmdBank 
                     Height          =   315
                     Left            =   3675
                     TabIndex        =   70
                     ToolTipText     =   "Browse for checked in guests."
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   360
                     _ExtentX        =   635
                     _ExtentY        =   556
                     Caption         =   "..."
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
                     cBack           =   14737632
                  End
                  Begin TrueOleDBList80.TDBCombo tdbTmp3 
                     Bindings        =   "frmMDEmployee.frx":178F
                     DataMember      =   "tdbJob"
                     Height          =   300
                     Left            =   0
                     TabIndex        =   155
                     Top             =   615
                     Visible         =   0   'False
                     Width           =   2835
                     _ExtentX        =   5001
                     _ExtentY        =   529
                     _LayoutType     =   4
                     _RowHeight      =   -2147483647
                     _WasPersistedAsPixels=   0
                     _DropdownWidth  =   0
                     _EDITHEIGHT     =   529
                     _GAPHEIGHT      =   53
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns(1)._VlistStyle=   0
                     Columns(1)._MaxComboItems=   5
                     Columns(1).DataField=   ""
                     Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   2
                     Splits(0)._UserFlags=   0
                     Splits(0).ExtendRightColumn=   -1  'True
                     Splits(0).AllowRowSizing=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                     Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                     Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                     Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                     Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                     Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                     Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                     Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                     Splits.Count    =   1
                     Appearance      =   0
                     BorderStyle     =   1
                     ComboStyle      =   0
                     AutoCompletion  =   0   'False
                     LimitToList     =   0   'False
                     ColumnHeaders   =   -1  'True
                     ColumnFooters   =   0   'False
                     DataMode        =   0
                     DefColWidth     =   0
                     Enabled         =   -1  'True
                     HeadLines       =   0
                     FootLines       =   1
                     RowDividerStyle =   0
                     Caption         =   ""
                     EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                     LayoutName      =   ""
                     LayoutFileName  =   ""
                     MultipleLines   =   0
                     EmptyRows       =   -1  'True
                     CellTips        =   0
                     AutoSize        =   0   'False
                     ListField       =   ""
                     BoundColumn     =   ""
                     IntegralHeight  =   0   'False
                     CellTipsWidth   =   0
                     CellTipsDelay   =   1000
                     AutoDropdown    =   0   'False
                     RowTracking     =   -1  'True
                     RightToLeft     =   0   'False
                     RowMember       =   ""
                     MouseIcon       =   0
                     MouseIcon.vt    =   3
                     MousePointer    =   0
                     MatchEntryTimeout=   2000
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     AnimateWindow   =   0
                     AnimateWindowDirection=   0
                     AnimateWindowTime=   200
                     AnimateWindowClose=   0
                     DropdownPosition=   0
                     Locked          =   0   'False
                     ScrollTrack     =   0   'False
                     RowDividerColor =   14215660
                     RowSubDividerColor=   14215660
                     AddItemSeparator=   ";"
                     _PropDict       =   $"frmMDEmployee.frx":17A0
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                     _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
                  Begin VB.Label Label4 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "DEFAULT BANK ACCOUNT"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Index           =   4
                     Left            =   45
                     TabIndex        =   162
                     Top             =   75
                     Width           =   2085
                  End
                  Begin VB.Label Label3 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bank"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   255
                     Left            =   75
                     TabIndex        =   112
                     Top             =   360
                     Width           =   1290
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Acct No"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     Left            =   4230
                     TabIndex        =   111
                     Top             =   360
                     Width           =   1635
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3360
            Left            =   -13785
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   300
            Width           =   12870
            _cx             =   22701
            _cy             =   5927
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14737632
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               Height          =   3390
               Left            =   30
               TabIndex        =   115
               Top             =   -60
               Width           =   12810
               Begin TDBText6Ctl.TDBText txtTelno 
                  Height          =   300
                  Left            =   4605
                  TabIndex        =   18
                  Top             =   960
                  Width           =   1770
                  _Version        =   65536
                  _ExtentX        =   3122
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":184A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":18B6
                  Key             =   "frmMDEmployee.frx":18D4
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtMobileno 
                  Height          =   300
                  Left            =   6420
                  TabIndex        =   19
                  Top             =   960
                  Width           =   1770
                  _Version        =   65536
                  _ExtentX        =   3122
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1918
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1984
                  Key             =   "frmMDEmployee.frx":19A2
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtEmail 
                  Height          =   300
                  Left            =   8220
                  TabIndex        =   20
                  Top             =   960
                  Width           =   2205
                  _Version        =   65536
                  _ExtentX        =   3889
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":19E6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1A52
                  Key             =   "frmMDEmployee.frx":1A70
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   50
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
               Begin TDBText6Ctl.TDBText txtStreet 
                  Height          =   300
                  Left            =   4605
                  TabIndex        =   17
                  Top             =   360
                  Width           =   5820
                  _Version        =   65536
                  _ExtentX        =   10266
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1AB4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1B20
                  Key             =   "frmMDEmployee.frx":1B3E
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   100
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
               Begin TDBText6Ctl.TDBText txtEmrgncyName 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   26
                  Top             =   2985
                  Width           =   3420
                  _Version        =   65536
                  _ExtentX        =   6032
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1B82
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1BEE
                  Key             =   "frmMDEmployee.frx":1C0C
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   70
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
               Begin TDBText6Ctl.TDBText txtEmrgncyNo 
                  Height          =   300
                  Left            =   3630
                  TabIndex        =   27
                  Top             =   2985
                  Width           =   2865
                  _Version        =   65536
                  _ExtentX        =   5054
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1C50
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1CBC
                  Key             =   "frmMDEmployee.frx":1CDA
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   50
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
               Begin TDBText6Ctl.TDBText txtEmrgncyEmail 
                  Height          =   300
                  Left            =   6540
                  TabIndex        =   28
                  Top             =   2985
                  Width           =   6195
                  _Version        =   65536
                  _ExtentX        =   10927
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1D1E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1D8A
                  Key             =   "frmMDEmployee.frx":1DA8
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   100
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
               Begin TDBDate6Ctl.TDBDate txtBirthDate 
                  Height          =   300
                  Left            =   8865
                  TabIndex        =   23
                  Top             =   1500
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":1DEC
                  Caption         =   "frmMDEmployee.frx":1EF2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":1F58
                  Keys            =   "frmMDEmployee.frx":1F76
                  Spin            =   "frmMDEmployee.frx":1FD4
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
                  ForeColor       =   4210752
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
                  Text            =   "01/21/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39468
                  CenturyMode     =   0
               End
               Begin TDBText6Ctl.TDBText txtBirthPlace 
                  Height          =   300
                  Left            =   4605
                  TabIndex        =   24
                  Top             =   2055
                  Width           =   4380
                  _Version        =   65536
                  _ExtentX        =   7726
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":1FFC
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":2068
                  Key             =   "frmMDEmployee.frx":2086
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   100
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
               Begin TrueOleDBList80.TDBCombo tdbGender 
                  Bindings        =   "frmMDEmployee.frx":20CA
                  DataMember      =   "tdbJob"
                  Height          =   300
                  Left            =   4620
                  TabIndex        =   21
                  Top             =   1500
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":20DB
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin TrueOleDBList80.TDBCombo tdbCivilStatus 
                  Bindings        =   "frmMDEmployee.frx":2185
                  DataMember      =   "tdbJob"
                  Height          =   300
                  Left            =   6435
                  TabIndex        =   22
                  Top             =   1500
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":2196
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin lvButton.lvButtons_H cmdEducationalBackground 
                  Height          =   720
                  Left            =   10635
                  TabIndex        =   116
                  Top             =   1635
                  Width           =   2070
                  _ExtentX        =   3651
                  _ExtentY        =   1270
                  Caption         =   "Educational Background"
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin TDBText6Ctl.TDBText txtBloodType 
                  Height          =   300
                  Left            =   9015
                  TabIndex        =   25
                  Top             =   2055
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":2240
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":22AC
                  Key             =   "frmMDEmployee.frx":22CA
                  BackColor       =   -2147483643
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   10
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
               Begin lvButton.lvButtons_H cmdUploadPhoto 
                  Height          =   345
                  Left            =   1980
                  TabIndex        =   163
                  ToolTipText     =   "Upload Picture"
                  Top             =   2010
                  Width           =   285
                  _ExtentX        =   503
                  _ExtentY        =   609
                  Caption         =   "..."
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E0E0E0&
                  Height          =   150
                  Left            =   15
                  TabIndex        =   117
                  Top             =   2370
                  Width           =   12765
               End
               Begin lvButton.lvButtons_H cmdUploadSignature 
                  Height          =   345
                  Left            =   4095
                  TabIndex        =   164
                  ToolTipText     =   "Upload Signature"
                  Top             =   840
                  Width           =   285
                  _ExtentX        =   503
                  _ExtentY        =   609
                  Caption         =   "..."
                  CapAlign        =   2
                  BackStyle       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFore           =   33023
                  cFHover         =   33023
                  cGradient       =   0
                  Mode            =   0
                  Value           =   0   'False
                  cBack           =   -2147483633
               End
               Begin VB.Image imgSignature 
                  Appearance      =   0  'Flat
                  BorderStyle     =   1  'Fixed Single
                  Height          =   1050
                  Left            =   2295
                  Stretch         =   -1  'True
                  Top             =   150
                  Width           =   2100
               End
               Begin VB.Image imgPhoto 
                  Appearance      =   0  'Flat
                  BorderStyle     =   1  'Fixed Single
                  Height          =   2220
                  Left            =   60
                  Stretch         =   -1  'True
                  Top             =   150
                  Width           =   2220
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Blood Type"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   31
                  Left            =   9030
                  TabIndex        =   157
                  Top             =   1830
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   14
                  Left            =   4620
                  TabIndex        =   129
                  Top             =   135
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Telephone No"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   19
                  Left            =   4620
                  TabIndex        =   128
                  Top             =   720
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mobile No"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   20
                  Left            =   6435
                  TabIndex        =   127
                  Top             =   720
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Email Address"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   21
                  Left            =   8250
                  TabIndex        =   126
                  Top             =   720
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gender"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   22
                  Left            =   4620
                  TabIndex        =   125
                  Top             =   1275
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Civil Status"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   124
                  Top             =   1275
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Birth date"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   24
                  Left            =   8880
                  TabIndex        =   123
                  Top             =   1275
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Birth place"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   25
                  Left            =   4620
                  TabIndex        =   122
                  Top             =   1830
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "CONTACT PERSON IN CASE OF EMERGENCY:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   26
                  Left            =   90
                  TabIndex        =   121
                  Top             =   2535
                  Width           =   5130
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contact No."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   27
                  Left            =   3645
                  TabIndex        =   120
                  Top             =   2790
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   28
                  Left            =   6540
                  TabIndex        =   119
                  Top             =   2790
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   29
                  Left            =   180
                  TabIndex        =   118
                  Top             =   2790
                  Width           =   1725
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3360
            Left            =   -14085
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   300
            Width           =   12870
            _cx             =   22701
            _cy             =   5927
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14737632
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
            Begin VB.Frame frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   3390
               Index           =   0
               Left            =   30
               TabIndex        =   131
               Top             =   -60
               Width           =   12810
               Begin VB.CommandButton Command1 
                  Caption         =   "Command1"
                  Height          =   270
                  Left            =   7680
                  TabIndex        =   170
                  Top             =   375
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.CheckBox chkConfidential 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Confidential"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   225
                  Left            =   5670
                  TabIndex        =   165
                  Top             =   390
                  Width           =   1545
               End
               Begin TDBText6Ctl.TDBText txtEmpNo 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   0
                  Top             =   345
                  Width           =   2700
                  _Version        =   65536
                  _ExtentX        =   4762
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":230E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":237A
                  Key             =   "frmMDEmployee.frx":2398
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   10
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
               Begin TDBText6Ctl.TDBText txtLastname 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   2
                  Top             =   885
                  Width           =   2700
                  _Version        =   65536
                  _ExtentX        =   4762
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":23DC
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":2448
                  Key             =   "frmMDEmployee.frx":2466
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtFirstname 
                  Height          =   300
                  Left            =   2880
                  TabIndex        =   3
                  Top             =   885
                  Width           =   3330
                  _Version        =   65536
                  _ExtentX        =   5874
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":24AA
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":2516
                  Key             =   "frmMDEmployee.frx":2534
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtMiddleName 
                  Height          =   300
                  Left            =   6225
                  TabIndex        =   4
                  Top             =   885
                  Width           =   2715
                  _Version        =   65536
                  _ExtentX        =   4789
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":2578
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":25E4
                  Key             =   "frmMDEmployee.frx":2602
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TrueOleDBList80.TDBCombo tdbEmpStat 
                  Height          =   300
                  Left            =   4575
                  TabIndex        =   10
                  Top             =   2445
                  Width           =   4005
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":2646
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin TDBText6Ctl.TDBText txtBranch 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   5
                  Top             =   1395
                  Width           =   4005
                  _Version        =   65536
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":26F0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":275C
                  Key             =   "frmMDEmployee.frx":277A
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtDivision 
                  Height          =   300
                  Left            =   4575
                  TabIndex        =   6
                  Top             =   1395
                  Width           =   4005
                  _Version        =   65536
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":27BE
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":282A
                  Key             =   "frmMDEmployee.frx":2848
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBText6Ctl.TDBText txtCostCenter 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   7
                  Top             =   1920
                  Width           =   4005
                  _Version        =   65536
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":288C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":28F8
                  Key             =   "frmMDEmployee.frx":2916
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin lvButton.lvButtons_H cmdBranch 
                  Height          =   315
                  Left            =   4185
                  TabIndex        =   64
                  Top             =   1395
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   556
                  Caption         =   "..."
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
                  cBack           =   14737632
               End
               Begin lvButton.lvButtons_H cmdDivision 
                  Height          =   315
                  Left            =   8595
                  TabIndex        =   65
                  Top             =   1395
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   556
                  Caption         =   "..."
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
                  cBack           =   14737632
               End
               Begin lvButton.lvButtons_H cmdCostCenter 
                  Height          =   315
                  Left            =   4185
                  TabIndex        =   66
                  Top             =   1920
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   556
                  Caption         =   "..."
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
                  cBack           =   14737632
               End
               Begin lvButton.lvButtons_H cmdSection 
                  Height          =   315
                  Left            =   8595
                  TabIndex        =   67
                  Top             =   1920
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   556
                  Caption         =   "..."
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
                  cBack           =   14737632
               End
               Begin TDBText6Ctl.TDBText txtSection 
                  Height          =   300
                  Left            =   4575
                  TabIndex        =   8
                  Top             =   1920
                  Width           =   4005
                  _Version        =   65536
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":295A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":29C6
                  Key             =   "frmMDEmployee.frx":29E4
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin TDBDate6Ctl.TDBDate txtDateHired 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   11
                  Top             =   2985
                  Width           =   1125
                  _Version        =   65536
                  _ExtentX        =   1984
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":2A28
                  Caption         =   "frmMDEmployee.frx":2B2E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":2B94
                  Keys            =   "frmMDEmployee.frx":2BB2
                  Spin            =   "frmMDEmployee.frx":2C10
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   4210752
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
                  Text            =   "09/29/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39720
                  CenturyMode     =   0
               End
               Begin TrueOleDBList80.TDBCombo tdbIsActive 
                  Bindings        =   "frmMDEmployee.frx":2C38
                  DataMember      =   "tdbJob"
                  Height          =   300
                  Left            =   1350
                  TabIndex        =   12
                  Top             =   2985
                  Width           =   3180
                  _ExtentX        =   5609
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":2C49
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin TDBDate6Ctl.TDBDate txtDateSuspended 
                  Height          =   300
                  Left            =   4575
                  TabIndex        =   13
                  Top             =   2985
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":2CF3
                  Caption         =   "frmMDEmployee.frx":2DF9
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":2E5F
                  Keys            =   "frmMDEmployee.frx":2E7D
                  Spin            =   "frmMDEmployee.frx":2EDB
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   4210752
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
                  Text            =   "09/29/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39720
                  CenturyMode     =   0
               End
               Begin TDBDate6Ctl.TDBDate txtDateResigned 
                  Height          =   300
                  Left            =   5670
                  TabIndex        =   14
                  Top             =   2985
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":2F03
                  Caption         =   "frmMDEmployee.frx":3009
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":306F
                  Keys            =   "frmMDEmployee.frx":308D
                  Spin            =   "frmMDEmployee.frx":30EB
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   4210752
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
                  Text            =   "09/29/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39720
                  CenturyMode     =   0
               End
               Begin TDBDate6Ctl.TDBDate txtDateProby 
                  Height          =   300
                  Left            =   6765
                  TabIndex        =   15
                  Top             =   2985
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":3113
                  Caption         =   "frmMDEmployee.frx":3219
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":327F
                  Keys            =   "frmMDEmployee.frx":329D
                  Spin            =   "frmMDEmployee.frx":32FB
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   4210752
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
                  Text            =   "09/29/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39720
                  CenturyMode     =   0
               End
               Begin TDBDate6Ctl.TDBDate txtDateRegularized 
                  Height          =   300
                  Left            =   7860
                  TabIndex        =   16
                  Top             =   2985
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   529
                  Calendar        =   "frmMDEmployee.frx":3323
                  Caption         =   "frmMDEmployee.frx":3429
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":348F
                  Keys            =   "frmMDEmployee.frx":34AD
                  Spin            =   "frmMDEmployee.frx":350B
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   4210752
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
                  Text            =   "09/29/2008"
                  ValidateMode    =   0
                  ValueVT         =   7
                  Value           =   39720
                  CenturyMode     =   0
               End
               Begin TDBText6Ctl.TDBText txtJobtitle 
                  Height          =   300
                  Left            =   165
                  TabIndex        =   9
                  Top             =   2445
                  Width           =   4005
                  _Version        =   65536
                  _ExtentX        =   7064
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":3533
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":359F
                  Key             =   "frmMDEmployee.frx":35BD
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin lvButton.lvButtons_H cmdJobTitle 
                  Height          =   315
                  Left            =   4185
                  TabIndex        =   68
                  Top             =   2445
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   556
                  Caption         =   "..."
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
                  cBack           =   14737632
               End
               Begin TrueOleDBList80.TDBCombo tdbTmp 
                  Bindings        =   "frmMDEmployee.frx":3601
                  DataMember      =   "tdbJob"
                  Height          =   300
                  Left            =   9885
                  TabIndex        =   153
                  Top             =   2670
                  Visible         =   0   'False
                  Width           =   2835
                  _ExtentX        =   5001
                  _ExtentY        =   529
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  _DropdownWidth  =   0
                  _EDITHEIGHT     =   529
                  _GAPHEIGHT      =   53
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).DataField=   ""
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).DataField=   ""
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   2
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=2"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits.Count    =   1
                  Appearance      =   0
                  BorderStyle     =   1
                  ComboStyle      =   0
                  AutoCompletion  =   0   'False
                  LimitToList     =   0   'False
                  ColumnHeaders   =   -1  'True
                  ColumnFooters   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   0
                  FootLines       =   1
                  RowDividerStyle =   0
                  Caption         =   ""
                  EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTips        =   0
                  AutoSize        =   0   'False
                  ListField       =   ""
                  BoundColumn     =   ""
                  IntegralHeight  =   0   'False
                  CellTipsWidth   =   0
                  CellTipsDelay   =   1000
                  AutoDropdown    =   0   'False
                  RowTracking     =   -1  'True
                  RightToLeft     =   0   'False
                  RowMember       =   ""
                  MouseIcon       =   0
                  MouseIcon.vt    =   3
                  MousePointer    =   0
                  MatchEntryTimeout=   2000
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  AnimateWindow   =   0
                  AnimateWindowDirection=   0
                  AnimateWindowTime=   200
                  AnimateWindowClose=   0
                  DropdownPosition=   0
                  Locked          =   0   'False
                  ScrollTrack     =   0   'False
                  RowDividerColor =   14215660
                  RowSubDividerColor=   14215660
                  AddItemSeparator=   ";"
                  _PropDict       =   $"frmMDEmployee.frx":3612
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
                  _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
               Begin TDBText6Ctl.TDBText txtIDNo 
                  Height          =   300
                  Left            =   2895
                  TabIndex        =   1
                  Top             =   345
                  Width           =   2700
                  _Version        =   65536
                  _ExtentX        =   4762
                  _ExtentY        =   529
                  Caption         =   "frmMDEmployee.frx":36BC
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployee.frx":3728
                  Key             =   "frmMDEmployee.frx":3746
                  BackColor       =   16777215
                  EditMode        =   0
                  ForeColor       =   4210752
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
                  MaxLength       =   20
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
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ID No"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   30
                  Left            =   2910
                  TabIndex        =   156
                  Top             =   135
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Employee No"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   0
                  Left            =   180
                  TabIndex        =   147
                  Top             =   120
                  Width           =   1725
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Last Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   0
                  Left            =   180
                  TabIndex        =   146
                  Top             =   675
                  Width           =   1725
               End
               Begin VB.Label label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "First Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   1
                  Left            =   2085
                  TabIndex        =   145
                  Top             =   675
                  Width           =   1725
               End
               Begin VB.Label label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Middle Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   2
                  Left            =   5595
                  TabIndex        =   144
                  Top             =   675
                  Width           =   1725
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Branch"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   0
                  Left            =   180
                  TabIndex        =   143
                  Top             =   1185
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Division"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   3
                  Left            =   4575
                  TabIndex        =   142
                  Top             =   1185
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cost Center"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   4
                  Left            =   180
                  TabIndex        =   141
                  Top             =   1710
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Section"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   6
                  Left            =   4575
                  TabIndex        =   140
                  Top             =   1710
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Employee Type"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   7
                  Left            =   4575
                  TabIndex        =   139
                  Top             =   2235
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Job Title"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   8
                  Left            =   180
                  TabIndex        =   138
                  Top             =   2235
                  Width           =   1725
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Hired"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   12
                  Left            =   180
                  TabIndex        =   137
                  Top             =   2775
                  Width           =   1080
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Emp. Status"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   13
                  Left            =   1365
                  TabIndex        =   136
                  Top             =   2775
                  Width           =   3345
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suspended"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   15
                  Left            =   4575
                  TabIndex        =   135
                  Top             =   2775
                  Width           =   1080
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Resigned"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   16
                  Left            =   5670
                  TabIndex        =   134
                  Top             =   2775
                  Width           =   1080
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Proby"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   17
                  Left            =   6765
                  TabIndex        =   133
                  Top             =   2775
                  Width           =   1080
               End
               Begin VB.Label label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Regularized"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   210
                  Index           =   18
                  Left            =   7860
                  TabIndex        =   132
                  Top             =   2775
                  Width           =   1080
               End
            End
         End
      End
      Begin VB.Frame fraButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   -15
         TabIndex        =   152
         Top             =   4140
         Width           =   13035
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   0
            Left            =   60
            TabIndex        =   61
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&NEW"
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
            Image           =   "frmMDEmployee.frx":378A
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   1
            Left            =   1515
            TabIndex        =   60
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&SAVE"
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
            Image           =   "frmMDEmployee.frx":5464
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   2
            Left            =   2970
            TabIndex        =   62
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&DELETE"
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
            Image           =   "frmMDEmployee.frx":5BDE
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   3
            Left            =   4425
            TabIndex        =   63
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&CLOSE"
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
            Image           =   "frmMDEmployee.frx":78B8
            cBack           =   14737632
         End
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13065
      TabIndex        =   71
      Top             =   0
      Width           =   13065
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Master data - Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   315
         Index           =   5
         Left            =   135
         TabIndex        =   72
         Top             =   225
         Width           =   5445
      End
   End
   Begin MSComDlg.CommonDialog dlgBrowsePic 
      Left            =   630
      Top             =   5745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBrowseSig 
      Left            =   1080
      Top             =   5745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMDEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mSave                         As Boolean
Dim mEduOpened                    As Boolean
Dim mBankAcctOpened               As Boolean
Dim mNonTaxAllowOpened            As Boolean

Public mEmployeeCode              As String
Dim mPicName                      As String

Public rsBankAccount              As ADODB.Recordset
Public rsEducationalBackground    As ADODB.Recordset
Public rsNonTaxAllow              As ADODB.Recordset
Dim rsEmployee                    As ADODB.Recordset

Dim mTxt                          As TDBText

Private Sub cmdAddBankAccount_Click()

  Dim rsTmp       As ADODB.Recordset
  
  With frmAdBanks
  
    If Not mBankAcctOpened Then
      Create_BankAccountTmp
      mBankAcctOpened = True
      If mSave = False Then
        NetOpen rsTmp, "select a.*,b.bankname from bankaccount a " & _
                       "left outer join bank b on a.bankcode = b.bankcode where a.employeecode = " & mEmployeeCode & ""
        If rsTmp.RecordCount > 0 Then
          rsTmp.MoveFirst
          Do While Not rsTmp.EOF
            With rsBankAccount
              .AddNew
              .Fields("bankcode") = rsTmp!bankcode & ""
              .Fields("bankname") = rsTmp!Bankname & ""
              .Fields("bankacctno") = rsTmp!bankacctno & ""
            End With
            rsTmp.MoveNext
          Loop
        End If
      End If
    End If
    
    Set .tdgBankAccount.DataSource = rsBankAccount
    .Show vbModal
    
  End With
  
End Sub

Private Sub cmdBank_Click()
  
  bind_tdb ConMain, tdbTmp3, "select bankcode,bankname from bank order by bankname", "bankname", "bankcode"
  Set mTxt = txtBank
  tdbTmp3.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp3.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp3.Visible = True
  tdbTmp3.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdBranch_Click()
  
  bind_tdb ConMain, tdbTmp, "select branchcode,branch from branch order by branch", "branch", "branchcode"
  Set mTxt = txtBranch
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdCostCenter_Click()
  
  bind_tdb ConMain, tdbTmp, "select CostCentercode,CostCenter from CostCenter order by CostCenter", "CostCenter", "CostCentercode"
  Set mTxt = txtCostCenter
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdDivision_Click()

  bind_tdb ConMain, tdbTmp, "select divisioncode,division from division order by division", "division", "divisioncode"
  Set mTxt = txtDivision
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdEducationalBackground_Click()
  
  Dim rsTmp       As ADODB.Recordset
  
  With frmAdEducationalBackground
  
    If Not mEduOpened Then
      Create_EducationalBackGroundTmp
      mEduOpened = True
      If mSave = False Then
        NetOpen rsTmp, "select * from educationalbackground where employeecode = " & mEmployeeCode & ""
        If rsTmp.RecordCount > 0 Then
          rsTmp.MoveFirst
          Do While Not rsTmp.EOF
            With rsEducationalBackground
              .AddNew
              .Fields("schoolattended") = rsTmp!schoolattended & ""
              .Fields("schooladdress") = rsTmp!schooladdress & ""
              .Fields("schoollevel") = rsTmp!schoollevel & ""
              .Fields("coursedescription") = rsTmp!coursedescription & ""
              If Not IsNull(rsTmp!fromyear) Then
                .Fields("fromyear") = Format(rsTmp!fromyear, "MM/DD/YYYY") & ""
              End If
              If Not IsNull(rsTmp!toyear) Then
                .Fields("toyear") = Format(rsTmp!toyear, "MM/DD/YYYY") & ""
              End If
              .Update
            End With
            rsTmp.MoveNext
          Loop
        End If
      End If
    End If
    
    Set .tdgEducationalBackground.DataSource = rsEducationalBackground
    .Show vbModal
    
  End With
  
End Sub

Private Sub cmdJobTitle_Click()
  
  bind_tdb ConMain, tdbTmp, "select JobTitlecode,description JobTitle from JobTitle order by JobTitle", "JobTitle", "JobTitlecode"
  Set mTxt = txtJobTitle
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdmenu_Click(Index As Integer)
  
  Select Case Index
    Case 0:
            If Not mSave Then
              If MsgBox("Do you want to create a new employee?", vbQuestion + vbYesNo) = vbYes Then
                ClearText
              End If
            Else
              ClearText
            End If
    Case 1: Save_Update
    Case 2: Delete_Employee
    Case 3: Unload Me
  End Select
  
End Sub

Private Sub cmdNonTaxAllow_Click()
    
  Dim rsTmp       As ADODB.Recordset
  
  If Not mNonTaxAllowOpened Then
    Create_NonTaxAllowTmp
    mNonTaxAllowOpened = True
    
    NetOpen rsTmp, "SELECT x1.*,IFNULL(x2.nontaxallow_amt,0) nontaxallow_amt FROM nontaxallow x1 " & _
                    "LEFT OUTER JOIN (SELECT * FROM employee_nontaxallow " & _
                    "                 WHERE employeecode =CONVERT(CONCAT('0','" & mEmployeeCode & "'),UNSIGNED INTEGER)) x2 ON x1.nontaxallow_id=x2.nontaxallow_id"
                      
    If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
        With rsNonTaxAllow
          .AddNew
          .Fields("nontaxallow_id") = rsTmp!nontaxallow_id & ""
          .Fields("nontaxallow_description") = rsTmp!nontaxallow_description & ""
          .Fields("nontaxallow_amt") = rsTmp!nontaxallow_amt & ""
          .Update
        End With
        rsTmp.MoveNext
      Loop
    End If
  End If
  
  With frmAdNonTaxAllow
    Set .tdgNonTaxAllow.DataSource = rsNonTaxAllow
    .Show vbModal
  End With
  
End Sub

Private Sub cmdSection_Click()
  
  bind_tdb ConMain, tdbTmp, "select Sectioncode,Sectionname from Section order by Sectionname", "Sectionname", "Sectioncode"
  Set mTxt = txtSection
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdShowRegistration_Click()
  With frmBrowseEmployee
    .mBrowseType = "Employee Masterdata"
    .Show vbModal
  End With
End Sub

Private Sub cmdUploadPhoto_Click()
    
    Dim mFileName As String
    If Not IsNull(imgPhoto.Picture) Then
      mFileName = dlgBrowsePic.FileName
    End If
    
    dlgBrowsePic.FileName = "*.jpg;*.gif"
    dlgBrowsePic.Flags = cdlOFNFileMustExist
    dlgBrowsePic.DialogTitle = "Browse Picture"
    dlgBrowsePic.ShowOpen
    
    If dlgBrowsePic.FileName = "*.jpg;*.gif" Then
      If Trim(mFileName) <> "" Then
        imgPhoto.Picture = LoadPicture(mFileName)
      End If
      dlgBrowsePic.FileName = mFileName
    Else
      imgPhoto.Picture = LoadPicture(dlgBrowsePic.FileName)
    End If
    
End Sub

Private Sub cmdUploadSignature_Click()
    
    Dim mFileName As String
    
    If Not IsNull(imgSignature.Picture) Then
      mFileName = dlgBrowseSig.FileName
    End If
    
    dlgBrowseSig.FileName = "*.jpg;*.gif"
    dlgBrowseSig.Flags = cdlOFNFileMustExist
    dlgBrowseSig.DialogTitle = "Browse Picture"
    dlgBrowseSig.ShowOpen
    
    If dlgBrowseSig.FileName = "*.jpg;*.gif" Then
      If Trim(mFileName) <> "" Then
        imgSignature.Picture = LoadPicture(mFileName)
      End If
      dlgBrowseSig.FileName = mFileName
    Else
      imgSignature.Picture = LoadPicture(dlgBrowseSig.FileName)
    End If

End Sub

Private Sub cmdWT_Click()
    
  bind_tdb ConMain, tdbTmp2, "select wtcode,description from wt order by description", "description", "wtcode"
  Set mTxt = txtWT
  tdbTmp2.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp2.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp2.Visible = True
  tdbTmp2.SetFocus
  SendKeys "{F4}"

End Sub

Private Sub Command1_Click()

  Command1.Enabled = False
  Dim rsWage As New ADODB.Recordset
  NetOpen rsWage, "select * from a_wage"
  If rsWage.RecordCount > 0 Then
    rsWage.MoveFirst
    Do While Not rsWage.EOF
      ConMain.Execute "delete from employee_nontaxallow where employeecode = " & rsWage!employeecode
      ConMain.Execute "update employee set monthly_rate = " & rsWage!basic & " , daily_rate = " & rsWage!dailyrate & " , hourly_rate = " & rsWage!hourlyrate & " where employeecode = " & rsWage!employeecode & ""
      If CDbl(rsWage!rice) > 0 Then
        ConMain.Execute "insert into employee_nontaxallow (employeecode, nontaxallow_id,nontaxallow_amt) values " & _
                      "(" & rsWage!employeecode & ",1," & rsWage!rice & ")"
      End If
      rsWage.MoveNext
    Loop
  End If
  MsgBox "Process complete"
  Command1.Enabled = True
  
End Sub

Private Sub Form_Activate()

    Focus_MDIButton Me

End Sub

Private Sub Form_Load()
    
    Dim rsTmp         As ADODB.Recordset
    Dim i             As Integer
        
    Add_MDIButton Me.Name, Me.Tag
    
    Set mTxt = txtEmpNo
    mEmployeeCode = ""
    
    tabEmployee.CurrTab = 0
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 2
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "Male"
          Case 2: .Fields("description") = "Female"
        End Select
        .Update
      Next
    End With
    
    With tdbGender
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
          Case 1: .Fields("description") = "Single"
          Case 2: .Fields("description") = "Married"
          Case 3: .Fields("description") = "Widow"
          Case 4: .Fields("description") = "Widower"
          Case 5: .Fields("description") = "Separated"
        End Select
        .Update
      Next
    End With
    
    With tdbCivilStatus
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing

    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 7
        .AddNew
        Select Case i
          Case 1: .Fields("code") = "Y"
                  .Fields("description") = "Active"
          Case 2: .Fields("code") = "N"
                  .Fields("description") = "Inactive"
          Case 3: .Fields("code") = "VL"
                  .Fields("description") = "Vacation Leave"
          Case 4: .Fields("code") = "ML"
                  .Fields("description") = "Maternity Leave"
          Case 5: .Fields("code") = "SL"
                  .Fields("description") = "Sick Leave"
          Case 6: .Fields("code") = "AW"
                  .Fields("description") = "AWOL"
          Case 7: .Fields("code") = "SP"
                  .Fields("description") = "Suspended"
        End Select
        .Update
      Next
    End With
    
    With tdbIsActive
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    bind_tdb ConMain, tdbPayFrequency, "select payfreqcode,payfreqname from payfrequency order by payfreqname", "payfreqname", "payfreqcode"
    bind_tdb ConMain, tdbRateType, "select ratetypecode, ratetypename from ratetypes order by ratetypename", "ratetypename", "ratetypecode"
    bind_tdb ConMain, tdbWrkDays, "select wrkdays_id, ttl_days from wrkdays order by wrkdays_id", "ttl_days", "wrkdays_id"
    bind_tdb ConMain, tdbEmpStat, "select empstatcode,empstatname from employmentstatus order by empstatname", "empstatname", "empstatcode"
    
    ClearText
    
End Sub

Private Sub Form_Resize()
  
  On Error Resume Next
  
  With fraHolder
    .Top = ((Me.ScaleHeight - (pic1.Top + pic1.Height)) / 2) - (.Height / 2)
    .Left = (Me.ScaleWidth / 2) - (.Width / 2)
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Public Sub ClearText()
      
      mSave = True
      mEduOpened = False
      mBankAcctOpened = False
      mNonTaxAllowOpened = False
      Lock_Button "TTFT", cmdMenu, 3
      lblEmployeeName.Caption = ""
      mEmployeeCode = ""
      
      tdbTmp_LostFocus
      tdbTmp2_LostFocus
      tdbTmp3_LostFocus
          
      '1st tab
      txtEmpNo.Text = ""
      txtIDNo.Text = ""
      txtLastname.Text = ""
      txtFirstname.Text = ""
      txtMiddleName.Text = ""
      txtBranch.Text = ""
      txtBranch.Tag = ""
      txtDivision.Text = ""
      txtDivision.Tag = ""
      txtCostCenter.Text = ""
      txtCostCenter.Tag = ""
      txtSection.Text = ""
      txtSection.Tag = ""
      txtJobTitle.Text = ""
      txtJobTitle.Tag = ""
      tdbEmpStat.BoundText = ""
      txtDateHired.Text = ""
      tdbIsActive.BoundText = ""
      txtDateSuspended.Text = ""
      txtDateResigned.Text = ""
      txtDateProby.Text = ""
      txtDateRegularized.Text = ""
      imgPhoto.Picture = Nothing
      imgSignature.Picture = Nothing
      dlgBrowsePic.FileName = ""
      dlgBrowseSig.FileName = ""
      chkConfidential.Value = vbUnchecked
      
      '2nd tab
      txtStreet.Text = ""
      txtTelno.Text = ""
      txtMobileno.Text = ""
      txtEmail.Text = ""
      tdbGender.Text = ""
      tdbCivilStatus.Text = ""
      txtBirthDate.Text = ""
      txtBirthPlace.Text = ""
      txtBloodType.Text = ""
      txtEmrgncyName.Text = ""
      txtEmrgncyNo.Text = ""
      txtEmrgncyEmail.Text = ""
      
      '3rd tab
      txtTinno.Text = ""
      txtWT.Text = ""
      txtWT.Tag = ""
      optTaxAuto.Value = True
      txtTaxAmt.Text = "0.00"
      txtSSSno.Text = ""
      optSSSAuto.Value = True
      txtSSSAmt.Text = "0.00"
      txtSssEr.Text = "0.00"
      txtSssEc.Text = "0.00"
      txtHDMFNo.Text = ""
      optHDMFAuto.Value = True
      txtHdmfAmt.Text = "0.00"
      txtHDMFEr.Text = "0.00"
      txtPhilHNo.Text = ""
      optPhilHAuto.Value = True
      txtPhilHAmt.Text = "0.00"
      txtPhilEr.Text = "0.00"
      txtBank.Text = ""
      txtBank.Tag = ""
      txtBankAcctNo.Text = ""
      chkSalToBank.Value = 0
      
      '4th tab
      tdbRateType.BoundText = ""
      tdbWrkDays.BoundText = ""
      tdbPayFrequency.BoundText = ""
      txtMonthly_Rate.Text = "0.0000"
      txtDaily_Rate.Text = "0.0000000"
      txtHourly_Rate.Text = "0.0000000"
      txtMealAllow.Text = "0.00"
      txtFixedEarnings.Text = "0.00"
      txtCola.Text = "0.00"
      
      Set rsBankAccount = Nothing
      Set rsEducationalBackground = Nothing
      Set rsNonTaxAllow = Nothing

      
End Sub

Public Sub AssignValue()

'Public Sub AssignValue(mEmployeeCode As Integer)
      
      NetOpen rsEmployee, "select x1.*,x2.branch,x3.division,x4.costcenter,x5.sectionname,x6.description jobtitle,x7.description wt,x8.bankname,x9.filename,x10.filename as sigfilename from employee x1 " & _
                          "left outer join branch x2 on x1.branchcode = x2.branchcode " & _
                          "left outer join division x3 on x1.divisioncode = x3.divisioncode " & _
                          "left outer join costcenter x4 on x1.costcentercode = x4.costcentercode " & _
                          "left outer join section x5 on x1.sectioncode = x5.sectioncode " & _
                          "left outer join jobtitle x6 on x1.jobtitlecode = x6.jobtitlecode " & _
                          "left outer join wt x7 on x1.wtcode = x7.wtcode  " & _
                          "left outer join bank x8 on x1.bankcode = x8.bankcode " & _
                          "left outer join emppics x9 on x1.employeecode = x9.employeecode " & _
                          "left outer join empsig x10 on x1.employeecode = x10.employeecode " & _
                          "where x1.employeecode = " & mEmployeeCode & ""
                    
      With rsEmployee
        If .RecordCount > 0 Then
          If Not .EOF Then
            mSave = False
            Lock_Button "TTTT", cmdMenu, 3
            lblEmployeeName.Caption = !dummycode & "   -   " & !lastname & ", " & !firstname & " " & !middlename & ""
            '1st tab
            txtEmpNo.Text = !dummycode & ""
            txtIDNo.Text = !idno & ""
            txtLastname.Text = !lastname & ""
            txtFirstname.Text = !firstname & ""
            txtMiddleName.Text = !middlename & ""
            txtBranch.Text = !branch & ""
            txtBranch.Tag = !branchcode & ""
            txtDivision.Text = !Division & ""
            txtDivision.Tag = !divisioncode & ""
            txtCostCenter.Text = !CostCenter & ""
            txtCostCenter.Tag = !costcentercode & ""
            txtSection.Text = !SectionName & ""
            txtSection.Tag = !sectioncode & ""
            txtJobTitle.Text = !jobtitle & ""
            txtJobTitle.Tag = !jobtitlecode & ""
            tdbEmpStat.BoundText = !empstatcode & ""
            txtDateHired.Text = IIf(IsNull(!datehired), "", Format(!datehired, "MM/DD/YYYY"))
            tdbIsActive.BoundText = !isactive & ""
            txtDateSuspended.Text = IIf(IsNull(!DateSuspended), "", Format(!DateSuspended, "MM/DD/YYYY"))
            txtDateResigned.Text = IIf(IsNull(!dateresigned), "", Format(!dateresigned, "MM/DD/YYYY"))
            txtDateProby.Text = IIf(IsNull(!dateproby), "", Format(!dateproby, "MM/DD/YYYY"))
            txtDateRegularized.Text = IIf(IsNull(!dateregularized), "", Format(!dateregularized, "MM/DD/YYYY"))
            chkConfidential.Value = IIf(!confidential = "N", vbUnchecked, vbChecked)
            
            If Not Dir(App.Path & "\EmpPics", vbDirectory) = vbNullString Then
              If Not IsNull(!FileName) Then
                If Not Dir(App.Path & "\EmpPics\" & !FileName, vbNormal) = vbNullString Then
                    imgPhoto.Picture = LoadPicture(App.Path & "\EmpPics\" & !FileName)
                End If
              End If
            End If
            
            If Not Dir(App.Path & "\EmpSig", vbDirectory) = vbNullString Then
              If Not IsNull(!sigfilename) Then
                If Not Dir(App.Path & "\EmpSig\" & !sigfilename, vbNormal) = vbNullString Then
                    imgSignature.Picture = LoadPicture(App.Path & "\EmpSig\" & !sigfilename)
                End If
              End If
            End If
            
            '2nd tab
            txtStreet.Text = !street & ""
            txtTelno.Text = !telno & ""
            txtMobileno.Text = !mobileno & ""
            txtEmail.Text = !email & ""
            tdbGender.Text = !gender & ""
            tdbCivilStatus.Text = !CivilStatus & ""
            txtBirthDate.Text = IIf(IsNull(!birthdate), "", Format(!birthdate, "MM/DD/YYYY"))
            txtBirthPlace.Text = !birthplace & ""
            txtBloodType.Text = !bloodtype & ""
            txtEmrgncyName.Text = !EmrgncyName & ""
            txtEmrgncyNo.Text = !EmrgncyNo & ""
            txtEmrgncyEmail.Text = !EmrgncyEmail & ""
            
            '3rd tab
            txtTinno.Text = !tinno & ""
            txtWT.Text = !WT & ""
            txtWT.Tag = !wtcode & ""
            
            If !taxauto <> 0 Then
                optTaxAuto.Value = True
            Else
                optTaxFixed.Value = True
            End If
            
            txtTaxAmt.Text = Format(!taxamt, "#,##0.00")
            txtSSSno.Text = !sssno & ""
            
            If !sssauto <> 0 Then
                optSSSAuto.Value = True
            Else
                optSSSFixed.Value = True
            End If
            
            txtSSSAmt.Text = Format(!sssamt, "#,##0.00")
            txtSssEr.Text = Format(!SssEr, "#,##0.00")
            txtSssEc.Text = Format(!sssEc, "#,##0.00")
            txtHDMFNo.Text = !hdmfno & ""
            
            If !hdmfauto <> 0 Then
                optHDMFAuto.Value = True
            Else
                optHDMFFixed.Value = True
            End If
            
            txtHdmfAmt.Text = Format(!HdmfAmt, "#,##0.00")
            txtHDMFEr.Text = Format(!HDMFEr, "#,##0.00")
            txtPhilHNo.Text = !philhno & ""
            optPhilHAuto.Value = True
            If !philhauto <> 0 Then
                optPhilHAuto.Value = True
            Else
                optPhilHFixed.Value = True
            End If
            txtPhilHAmt.Text = Format(!PhilHAmt, "#,##0.00")
            txtPhilEr.Text = Format(!philher, "#,##0.00")
            txtBank.Text = !Bankname & ""
            txtBank.Tag = !bankcode & ""
            txtBankAcctNo.Text = !bankacctno & ""
            chkSalToBank.Value = IIf(!saltobank = "Y", 1, 0)
            
            '4th tab
            tdbRateType.BoundText = !ratetypecode & ""
            tdbWrkDays.BoundText = !wrkdays_id & ""
            tdbPayFrequency.BoundText = !payfreqcode & ""
            txtMonthly_Rate.Text = Format(!monthly_rate, "#,##0.0000")
            txtDaily_Rate.Text = Format(!daily_rate, "#,##0.0000000")
            txtHourly_Rate.Text = Format(!Hourly_Rate, "#,##0.0000000")
            txtMealAllow.Text = Format(!MealAllow, "#,##0.00")
            txtFixedEarnings.Text = Format(!FixedEarnings, "#,##0.00")
            txtCola.Text = Format(!cola, "#,##0.00")
        Else
          MsgBox "Employee not found.", vbExclamation + vbOKOnly
        End If
      Else
        MsgBox "Employee not found.", vbExclamation + vbOKOnly
      End If
    End With
    
End Sub

Private Sub Save_Update()

  Dim rsEmpPics               As ADODB.Recordset
  Dim rsEmpSig                As ADODB.Recordset
  
  Dim mPhoto                  As ADODB.Stream
  
  Dim mWT                     As String
  Dim mBranch                 As String
  Dim mDivision               As String
  Dim mCostCenter             As String
  Dim mSection                As String
  Dim mJobTitle               As String
  Dim mDateSuspended          As String
  Dim mDateResigned           As String
  Dim mDateProby              As String
  Dim mDateRegularized        As String
  Dim mBirthDate              As String
  Dim mBank                   As String
  

  Dim mSSSAuto                As Integer
  Dim mPHILHAuto              As Integer
  Dim mHDMFAuto               As Integer
  Dim mTAXAuto                As Integer
  Dim mLneNo                  As Integer
  
  Dim mSSSAmt                 As Double
  Dim mSSSEr                  As Double
  Dim mSSSEc                  As Double
  Dim mPHILHAmt               As Double
  Dim mPHILHEr                As Double
  Dim mHDMFAmt                As Double
  Dim mHdmfEr                 As Double
  Dim mTAXAmt                 As Double
  
  tdbTmp_LostFocus
  tdbTmp2_LostFocus
  tdbTmp3_LostFocus
  
    If Trim(txtLastname.Text) = "" Then
      MsgBox "Lastname is blank.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtLastname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtFirstname.Text) = "" Then
      MsgBox "Firstname is blank.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtFirstname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
      MsgBox "Middlename is blank.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtMiddleName.SetFocus
      Exit Sub
    End If
    
    If Trim(txtBranch.Tag) = "" Then
      MsgBox "Please select a branch.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtBranch.SetFocus
      Exit Sub
    Else
      mBranch = txtBranch.Tag
    End If
    
    If Trim(txtDivision.Tag) = "" Then
      MsgBox "Please select a division.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtDivision.SetFocus
      Exit Sub
    Else
      mDivision = txtDivision.Tag
    End If
    
    If Trim(txtCostCenter.Tag) = "" Then
      MsgBox "Please select a cost center.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtCostCenter.SetFocus
      Exit Sub
    Else
      mCostCenter = txtCostCenter.Tag
    End If
    
    If Trim(txtSection.Tag) <> "" Then
        mSection = txtSection.Tag
    Else
        mSection = "Null"
    End If
    
    If Trim(txtJobTitle.Tag) = "" Then
        MsgBox "Please select a job description.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 0
        txtJobTitle.SetFocus
        Exit Sub
    Else
      mJobTitle = txtJobTitle.Tag
    End If
        
    If Trim(tdbEmpStat.Text) = "" Or IsNull(tdbEmpStat.SelectedItem) Or tdbEmpStat.ApproxCount = 0 Then
        MsgBox "Please select an employmee type.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 0
        tdbEmpStat.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDateHired.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 0
      txtDateHired.SetFocus
    End If
    
    If Trim(tdbIsActive.Text) = "" Or IsNull(tdbIsActive.SelectedItem) Or tdbIsActive.ApproxCount = 0 Then
        MsgBox "Please select an employment status.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 0
        tdbIsActive.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDateSuspended.Text) Then
      mDateSuspended = "Null"
    Else
      mDateSuspended = "'" & Format(txtDateSuspended.Text, "YYYY-MM-DD") & "'"
    End If
    
    If Not IsDate(txtDateResigned.Text) Then
      mDateResigned = "Null"
    Else
      mDateResigned = "'" & Format(txtDateResigned.Text, "YYYY-MM-DD") & "'"
    End If
    
    If Not IsDate(txtDateProby.Text) Then
      mDateProby = "Null"
    Else
      mDateProby = "'" & Format(txtDateProby.Text, "YYYY-MM-DD") & "'"
    End If
    
    If Not IsDate(txtDateRegularized.Text) Then
      mDateRegularized = "Null"
    Else
      mDateRegularized = "'" & Format(txtDateRegularized.Text, "YYYY-MM-DD") & "'"
    End If
    
    If Trim(tdbGender.Text) = "" Or IsNull(tdbGender.SelectedItem) Or tdbGender.ApproxCount = 0 Then
      MsgBox "Please select a gender.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 1
      tdbGender.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbCivilStatus.Text) = "" Or IsNull(tdbCivilStatus.SelectedItem) Or tdbCivilStatus.ApproxCount = 0 Then
      MsgBox "Please select a civil status.", vbExclamation + vbOKOnly
      tabEmployee.CurrTab = 1
      tdbCivilStatus.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtBirthDate.Text) Then
      mBirthDate = "Null"
    Else
      mBirthDate = "'" & Format(txtBirthDate.Text, "YYYY-MM-DD") & "'"
    End If
    
    If optTaxAuto.Value = True Then
        mTAXAuto = 1
        mTAXAmt = 0
        If Trim(txtWT.Tag) = "" Then
            mWT = "Null"
        Else
            mWT = txtWT.Tag
        End If
    Else
        mTAXAuto = 0
        mTAXAmt = Format(txtTaxAmt.Text, "###0.00")
        mWT = "Null"
    End If
    
    If optSSSAuto.Value = True Then
        mSSSAuto = 1
        mSSSAmt = 0
        mSSSEr = 0
        mSSSEc = 0
    Else
        mSSSAuto = 0
        mSSSAmt = Format(txtSSSAmt.Text, "###0.00")
        mSSSEr = Format(txtSssEr.Text, "###0.00")
        mSSSEc = Format(txtSssEc.Text, "###0.00")
    End If
    
    If optHDMFAuto.Value = True Then
        mHDMFAuto = 1
    Else
        mHDMFAuto = 0
        mHDMFAmt = Format(txtHdmfAmt.Text, "###0.00")
        mHdmfEr = Format(txtHDMFEr.Text, "###0.00")
    End If
    
    If optPhilHAuto.Value = True Then
        mPHILHAuto = 1
        mPHILHAmt = 0
        mPHILHEr = 0
    Else
        mPHILHAuto = 0
        mPHILHAmt = Format(txtPhilHAmt.Text, "###0.00")
        mPHILHEr = Format(txtPhilEr.Text, "###0.00")
    End If

    If Trim(txtBank.Tag) = "" Then
      mBank = "Null"
    Else
      mBank = txtBank.Tag
    End If
    
    If Trim(tdbPayFrequency.Text) = "" And IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
        MsgBox "Please assign a payroll frequency.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 3
        tdbPayFrequency.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbRateType.Text) = "" Or IsNull(tdbRateType.SelectedItem) Or tdbRateType.ApproxCount = 0 Then
        MsgBox "Please select a rate type.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 3
        tdbRateType.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbWrkDays.Text) = "" Or IsNull(tdbWrkDays.SelectedItem) Or tdbWrkDays.ApproxCount = 0 Then
        MsgBox "Please select employee's number of workdays in a month.", vbExclamation + vbOKOnly
        tabEmployee.CurrTab = 3
        tdbWrkDays.SetFocus
        Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    
    If optTaxAuto.Value = False Then
        txtWT.Text = ""
        txtWT.Tag = ""
    Else
        txtTaxAmt.Text = "0.00"
    End If
    
    If optSSSAuto.Value = True Then
        txtSSSAmt.Text = "0.00"
        txtSssEr.Text = "0.00"
        txtSssEc.Text = "0.00"
    End If
    
    If optHDMFAuto.Value = True Then
        txtHdmfAmt.Text = "0.00"
        txtHDMFEr.Text = "0.00"
    End If
    
    If optPhilHAuto.Value = True Then
        txtPhilHAmt.Text = "0.00"
        txtPhilEr.Text = "0.00"
    End If
    
    If txtMonthly_Rate.Value > 0 Then
        If Trim(tdbWrkDays.Text) <> "" And Not IsNull(tdbWrkDays.SelectedItem) And tdbWrkDays.ApproxCount > 0 Then
            txtDaily_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text), "#,##0.0000000")
            txtHourly_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text) / 8, "#,##0.0000000")
        End If
    End If
    
    If mSave Then
      mEmployeeCode = LastCode("Employee")
    End If
    
    If dlgBrowsePic.FileName <> "" Then
    
      ConMain.Execute "delete from emppics where employeecode = " & mEmployeeCode & ""
      
      Set mPhoto = New ADODB.Stream
      mPhoto.Type = adTypeBinary
      mPhoto.Open
      mPhoto.LoadFromFile (dlgBrowsePic.FileName)
      
      NetOpen rsEmpPics, "select * from emppics limit 0"
      
      With rsEmpPics
        .AddNew
        .Fields("employeecode") = mEmployeeCode
        .Fields("images") = mPhoto.Read
        .Fields("filename") = mEmployeeCode & "." & GetFileExt(dlgBrowsePic.FileTitle)
        .Update
        mPicName = .Fields("filename")
      
        If Not Dir(App.Path & "\EmpPics", vbDirectory) = vbNullString Then
          
          If Not Dir(App.Path & "\EmpPics\" & !FileName, vbNormal) = vbNullString Then
              Kill (App.Path & "\EmpPics\" & !FileName)
          End If
          
          mPhoto.SaveToFile App.Path & "\EmpPics\" & !FileName
          
        End If
      End With
      
      Set mPhoto = Nothing
      Set rsEmpPics = Nothing
      
    End If
    
    If dlgBrowseSig.FileName <> "" Then
    
      ConMain.Execute "delete from empsig where employeecode = " & mEmployeeCode & ""
      
      Set mPhoto = New ADODB.Stream
      mPhoto.Type = adTypeBinary
      mPhoto.Open
      mPhoto.LoadFromFile (dlgBrowseSig.FileName)
      
      NetOpen rsEmpSig, "select * from empsig limit 0"
      With rsEmpSig
        .AddNew
        .Fields("employeecode") = mEmployeeCode
        .Fields("images") = mPhoto.Read
        .Fields("filename") = mEmployeeCode & "." & GetFileExt(dlgBrowseSig.FileTitle)
        .Update
        mPicName = .Fields("filename")
      
        If Not Dir(App.Path & "\empsig", vbDirectory) = vbNullString Then
          If Not Dir(App.Path & "\empsig\" & !FileName, vbNormal) = vbNullString Then
              Kill (App.Path & "\empsig\" & !FileName)
          End If
          mPhoto.SaveToFile App.Path & "\empsig\" & !FileName
        End If
      
      End With
      
      Set rsEmpSig = Nothing
      
    End If
    
    If mSave Then
    
      ConMain.Execute "insert into employee (employeecode,dummycode,biometid,idno,lastname,firstname,middlename,gender,civilstatus,birthdate,datehired, " & _
                          "datesuspended,dateresigned,dateproby,dateregularized,street,branchcode,divisioncode,costcentercode,sectioncode,telno,mobileno,email, " & _
                          "bloodtype,birthplace,emrgncyname,emrgncyno,emrgncyemail,payfreqcode,filename, " & _
                          "wtcode,jobtitlecode,empstatcode,ratetypecode,wrkdays_id,monthly_rate,daily_rate,hourly_rate, " & _
                          "sssno,philhno,tinno,hdmfno,bankcode,bankacctno, " & _
                          "sssamt,ssser,sssec,philhamt,philher,taxamt,hdmfamt,hdmfer, " & _
                          "sssauto,philhauto,taxauto,hdmfauto,saltobank," & _
                          "isactive,mealallow,fixedEarnings,cola,confidential) values " & _
                          "(" & mEmployeeCode & ",'" & Format(mEmployeeCode, "00000000") & "','" & mEmployeeCode & "','" & UCase(txtIDNo.Text) & "','" & UCase(txtLastname.Text) & "','" & UCase(txtFirstname.Text) & "','" & UCase(txtMiddleName.Text) & "','" & tdbGender.Text & "','" & tdbCivilStatus.Text & "'," & mBirthDate & ",'" & Format(txtDateHired.Text, "YYYY-MM-DD") & "'," & _
                          mDateSuspended & "," & mDateResigned & "," & mDateProby & "," & mDateRegularized & ",'" & Swap(txtStreet.Text) & "'," & mBranch & "," & mDivision & "," & mCostCenter & "," & mSection & ",'" & Swap(txtTelno.Text) & "','" & Swap(txtMobileno.Text) & "','" & Swap(txtEmail.Text) & "'," & _
                          "'" & Swap(txtBloodType.Text) & "','" & Swap(txtBirthPlace.Text) & "','" & Swap(UCase(txtEmrgncyName.Text)) & "','" & Swap(UCase(txtEmrgncyNo.Text)) & "','" & Swap(UCase(txtEmrgncyEmail.Text)) & "'," & tdbPayFrequency.BoundText & ",'" & mPicName & "', " & _
                          mWT & "," & mJobTitle & "," & tdbEmpStat.BoundText & "," & tdbRateType.BoundText & "," & tdbWrkDays.BoundText & "," & Format(txtMonthly_Rate.Text, "##0.00") & "," & Format(txtDaily_Rate.Text, "##0.0000000") & "," & Format(txtHourly_Rate.Text, "##0.0000000") & "," & _
                          "'" & Swap(txtSSSno.Text) & "','" & Swap(txtPhilHNo.Text) & "','" & Swap(txtTinno.Text) & "','" & Swap(txtHDMFNo.Text) & "'," & mBank & ",'" & Swap(txtBankAcctNo.Text) & "'," & _
                          mSSSAmt & "," & mSSSEr & "," & mSSSEc & "," & mPHILHAmt & "," & mPHILHEr & "," & mTAXAmt & "," & mHDMFAmt & "," & mHdmfEr & "," & _
                          mSSSAuto & "," & mPHILHAuto & "," & mTAXAuto & "," & mHDMFAuto & "," & IIf(chkSalToBank.Value <> 0, "'Y'", "'N'") & "," & _
                          "'" & tdbIsActive.BoundText & "'," & Format(txtMealAllow.Text, "##0.00") & "," & Format(txtFixedEarnings.Text, "##0.00") & "," & Format(txtCola.Text, "##0.00") & ",'" & IIf(chkConfidential.Value = vbChecked, "Y", "N") & "')"
    
      txtEmpNo.Text = Format(mEmployeeCode, "00000000")
      
      mSave = False
      
      MsgBox "New employee has been added.", vbInformation + vbOKOnly
      
  Else
  
      ConMain.Execute "update employee set idno = '" & UCase(txtIDNo.Text) & "',lastname = '" & UCase(txtLastname.Text) & "', firstname = '" & UCase(txtFirstname.Text) & "', middlename = '" & UCase(txtMiddleName.Text) & "', " & _
                          "gender = '" & tdbGender.Text & "', civilstatus = '" & tdbCivilStatus.Text & "', birthdate = " & mBirthDate & ", datehired = '" & Format(txtDateHired.Text, "YYYY-MM-DD") & "',street = '" & Swap(txtStreet.Text) & "',birthplace = '" & Swap(txtBirthPlace.Text) & "',bloodtype = '" & Swap(txtBloodType.Text) & "', " & _
                          "datesuspended=" & mDateSuspended & ",dateresigned=" & mDateResigned & ",dateproby= " & mDateProby & ",dateregularized = " & mDateRegularized & ",branchcode = " & mBranch & ", divisioncode = " & mDivision & ", costcentercode = " & mCostCenter & ",sectioncode = " & mSection & "," & _
                          "telno = '" & Swap(txtTelno.Text) & "', mobileno = '" & Swap(txtMobileno.Text) & "', email = '" & Swap(txtEmail.Text) & "', " & _
                          "emrgncyname = '" & Swap(UCase(txtEmrgncyName.Text)) & "', emrgncyno = '" & Swap(UCase(txtEmrgncyNo.Text)) & "', emrgncyemail = '" & Swap(UCase(txtEmrgncyEmail.Text)) & "',payfreqcode = " & tdbPayFrequency.BoundText & ", filename = '" & mPicName & "', " & _
                          "wtcode = " & mWT & ",jobtitlecode = " & mJobTitle & ",empstatcode = " & tdbEmpStat.BoundText & ",ratetypecode = " & tdbRateType.BoundText & ",wrkdays_id=" & tdbWrkDays.BoundText & ",monthly_rate = " & Format(txtMonthly_Rate.Text, "##0.00") & ",daily_rate = " & Format(txtDaily_Rate.Text, "##0.0000000") & ",hourly_rate = " & Format(txtHourly_Rate.Text, "##0.0000000") & "," & _
                          "sssno = '" & Swap(txtSSSno.Text) & "', philhno = '" & Swap(txtPhilHNo.Text) & "',tinno = '" & Swap(txtTinno.Text) & "',hdmfno = '" & Swap(txtHDMFNo.Text) & "', bankacctno = '" & Swap(txtBankAcctNo.Text) & "', " & _
                          "ssser = " & mSSSEr & ",sssamt = " & mSSSAmt & ",sssec = " & mSSSEc & ",philhamt = " & mPHILHAmt & ",philher = " & mPHILHEr & ",taxamt = " & mTAXAmt & ",hdmfamt = " & mHDMFAmt & ",hdmfer = " & mHdmfEr & ", " & _
                          "sssauto = " & mSSSAuto & ",philhauto = " & mPHILHAuto & ",taxauto = " & mTAXAuto & ",hdmfauto = " & mHDMFAuto & ", " & _
                          "saltobank = '" & IIf(chkSalToBank.Value <> 0, "Y", "N") & "',bankcode = " & mBank & "," & _
                          "isactive = '" & tdbIsActive.BoundText & "',mealallow = " & Format(txtMealAllow.Text, "##0.00") & ",fixedearnings = " & Format(txtFixedEarnings.Text, "##0.00") & ",cola = " & Format(txtCola.Text, "##0.00") & ",confidential='" & IIf(chkConfidential.Value = vbChecked, "Y", "N") & "' " & _
                          "where employeecode = " & mEmployeeCode & ""
    
    ConMain.Execute "update payroll set saltobank = '" & IIf(chkSalToBank.Value <> 0, "Y", "N") & "' where employeecode = " & rsEmployee!employeecode & " and fnlz = 'N'"
    MsgBox "Employee record has been updated", vbInformation + vbOKOnly
  End If

  If mEduOpened = True Then
    
    ConMain.Execute "delete from educationalbackground where employeecode = " & mEmployeeCode & ""
    
    With rsEducationalBackground
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          ConMain.Execute "insert into educationalbackground(employeecode,schoolattended,schooladdress,schoollevel,coursedescription,fromyear,toyear) values ( " & _
                        "" & mEmployeeCode & ",'" & Swap(!schoolattended) & "','" & Swap(!schooladdress) & "','" & !schoollevel & "','" & Swap(!coursedescription) & "'," & _
                        "" & IIf(IsDate(!fromyear), "'" & Format(!fromyear, "YYYY-MM-DD") & "'", "Null") & ", " & _
                        "" & IIf(IsDate(!toyear), "'" & Format(!toyear, "YYYY-MM-DD") & "'", "Null") & ")"
          .MoveNext
        Loop
      End If
    End With
    
  End If
  
  If mBankAcctOpened = True Then
    
    mLneNo = 1
    ConMain.Execute "delete from bankaccount where employeecode = " & mEmployeeCode & ""
    
    With rsBankAccount
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          ConMain.Execute "insert into bankaccount(bankaccountlne,employeecode,bankcode,bankacctno) values ( " & _
                        "" & mLneNo & "," & mEmployeeCode & "," & !bankcode & ",'" & !bankacctno & "')"
          mLneNo = mLneNo + 1
          .MoveNext
        Loop
      End If
    End With
    
  End If
  
  If mNonTaxAllowOpened = True Then
    
    ConMain.Execute "delete from employee_nontaxallow where employeecode = " & mEmployeeCode & ""
    
    With rsNonTaxAllow
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          If !nontaxallow_amt > 0 Then
            ConMain.Execute "insert into employee_nontaxallow(employeecode,nontaxallow_id,nontaxallow_amt) values ( " & _
                            "" & mEmployeeCode & "," & !nontaxallow_id & "," & !nontaxallow_amt & ")"
          End If
          .MoveNext
        Loop
      End If
    End With
    
  End If
                          
  ConMain.CommitTrans
  
  lblEmployeeName.Caption = txtEmpNo.Text & "   -   " & txtLastname.Text & ", " & txtFirstname.Text & " " & txtMiddleName.Text & ""
  dlgBrowsePic.FileName = ""
    
  
End Sub

Private Sub chkSalToBank_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tabEmployee.CurrTab = 3
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbTmp_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Trim(tdbTmp.Text) <> "" And Not IsNull(tdbTmp.SelectedItem) And tdbTmp.ApproxCount > 0 Then
      mTxt.Tag = tdbTmp.BoundText
      mTxt.Text = tdbTmp.Text
    Else
      mTxt.Tag = ""
      mTxt.Text = ""
    End If
    mTxt.Visible = True
    mTxt.SetFocus
    tdbTmp.Visible = False
  Else
    SearchList KeyAscii, tdbTmp, tdbTmp.RowSource, tdbTmp.Text
  End If
End Sub

Private Sub tdbTmp_LostFocus()
  mTxt.Visible = True
  tdbTmp.Visible = False
End Sub

Private Sub tdbTmp2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Trim(tdbTmp2.Text) <> "" And Not IsNull(tdbTmp2.SelectedItem) And tdbTmp2.ApproxCount > 0 Then
      mTxt.Tag = tdbTmp2.BoundText
      mTxt.Text = tdbTmp2.Text
    Else
      mTxt.Tag = ""
      mTxt.Text = ""
    End If
    mTxt.Visible = True
    mTxt.SetFocus
    tdbTmp2.Visible = False
  Else
    SearchList KeyAscii, tdbTmp2, tdbTmp2.RowSource, tdbTmp2.Text
  End If
End Sub

Private Sub tdbTmp2_LostFocus()
  mTxt.Visible = True
  tdbTmp2.Visible = False
End Sub

Private Sub tdbTmp3_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    If Trim(tdbTmp3.Text) <> "" And Not IsNull(tdbTmp3.SelectedItem) And tdbTmp3.ApproxCount > 0 Then
      mTxt.Tag = tdbTmp3.BoundText
      mTxt.Text = tdbTmp3.Text
    Else
      mTxt.Tag = ""
      mTxt.Text = ""
    End If
    mTxt.Visible = True
    mTxt.SetFocus
    tdbTmp3.Visible = False
  Else
    SearchList KeyAscii, tdbTmp3, tdbTmp3.RowSource, tdbTmp3.Text
  End If

End Sub

Public Sub Create_EducationalBackGroundTmp()

    Set rsEducationalBackground = Nothing
    Set rsEducationalBackground = New ADODB.Recordset
    
    With rsEducationalBackground
        .Fields.Append "schoolattended", adVarChar, 150
        .Fields.Append "schooladdress", adVarChar, 150
        .Fields.Append "schoollevel", adVarChar, 150
        .Fields.Append "coursedescription", adVarChar, 150
        .Fields.Append "fromyear", adDate
        .Fields.Append "toyear", adDate
        .Open
    End With
    
End Sub

Public Sub Create_NonTaxAllowTmp()

    Set rsNonTaxAllow = Nothing
    Set rsNonTaxAllow = New ADODB.Recordset
    
    With rsNonTaxAllow
        .Fields.Append "nontaxallow_id", adInteger
        .Fields.Append "nontaxallow_description", adVarChar, 50
        .Fields.Append "nontaxallow_amt", adDouble
        .Open
    End With
    
End Sub
Public Sub Create_BankAccountTmp()

    Set rsBankAccount = Nothing
    Set rsBankAccount = New ADODB.Recordset
    
    With rsBankAccount
        .Fields.Append "bankcode", adInteger
        .Fields.Append "bankname", adVarChar, 50
        .Fields.Append "bankacctno", adVarChar, 20
        .Open
    End With
    
End Sub


Private Sub tdbTmp3_LostFocus()
  mTxt.Visible = True
  tdbTmp3.Visible = False
End Sub

Private Sub tdbWrkDays_LostFocus()
    If txtMonthly_Rate.Value > 0 Then
        If Trim(tdbWrkDays.Text) <> "" And Not IsNull(tdbWrkDays.SelectedItem) And tdbWrkDays.ApproxCount > 0 Then
            txtDaily_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text), "#,##0.0000000")
            txtHourly_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text) / 8, "#,##0.0000000")
        End If
    End If
End Sub

Private Sub txtBirthDate_GotFocus()
  With txtBirthDate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbCivilStatus_GotFocus()
  With tdbCivilStatus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbCivilStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbEmpStat_GotFocus()
  With tdbEmpStat
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbEmpStat_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbGender_GotFocus()
  With tdbGender
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbGender_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbIsActive_GotFocus()
  With tdbIsActive
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbIsActive_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbPayFrequency_GotFocus()
  With tdbPayFrequency
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbPayFrequency_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbRateType_GotFocus()
  With tdbRateType
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbRateType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub tdbWrkDays_GotFocus()
    With tdbWrkDays
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tdbWrkDays_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtBank_GotFocus()
  With txtBank
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBank_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtBankAcctNo_GotFocus()
  With txtBankAcctNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBankAcctNo_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtBirthPlace_GotFocus()
  With txtBirthPlace
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBirthPlace_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtBloodType_GotFocus()
  With txtBloodType
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBloodType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtBranch_GotFocus()
  With txtBranch
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtBranch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdBranch_Click
End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtCola_GotFocus()
  With txtCola
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCola_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtCostCenter_GotFocus()
  With txtCostCenter
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCostCenter_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdCostCenter_Click
End Sub

Private Sub txtCostCenter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDaily_Rate_GotFocus()
  With txtDaily_Rate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDaily_Rate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateHired_GotFocus()
  With txtDateHired
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateHired_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateProby_GotFocus()
  With txtDateProby
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateProby_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateRegularized_GotFocus()
  With txtDateRegularized
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateRegularized_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tabEmployee.CurrTab = 1
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateResigned_GotFocus()
  With txtDateResigned
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateResigned_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateSuspended_GotFocus()
  With txtDateSuspended
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateSuspended_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDivision_GotFocus()
  With txtDivision
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDivision_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdDivision_Click
End Sub

Private Sub txtDivision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmail_GotFocus()
  With txtEmail
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmpNo_GotFocus()
  With txtEmpNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmpNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmrgncyEmail_GotFocus()
  With txtEmrgncyEmail
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmrgncyEmail_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tabEmployee.CurrTab = 2
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmrgncyName_GotFocus()
  With txtEmrgncyName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmrgncyName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmrgncyNo_GotFocus()
  With txtEmrgncyNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmrgncyNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtFirstname_GotFocus()
  With txtFirstname
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtFixedEarnings_GotFocus()
  With txtFixedEarnings
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtFixedEarnings_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtHdmfAmt_GotFocus()
  With txtHdmfAmt
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtHdmfAmt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtHDMFEr_GotFocus()
  With txtHDMFEr
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtHDMFEr_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtHDMFNo_GotFocus()
  With txtHDMFNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtHDMFNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtHourly_Rate_GotFocus()
  With txtHourly_Rate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtHourly_Rate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtIDNo_GotFocus()
  With txtIDNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtIDNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtJobTitle_GotFocus()
  With txtJobTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtJobTitle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdJobTitle_Click
End Sub

Private Sub txtJobTitle_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtLastname_GotFocus()
  With txtLastname
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMealAllow_GotFocus()
  With txtMealAllow
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtMealAllow_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMiddleName_GotFocus()
  With txtMiddleName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMobileno_GotFocus()
  With txtMobileno
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtMobileno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMonthly_Rate_GotFocus()
  With txtMonthly_Rate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtMonthly_Rate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtMonthly_Rate_LostFocus()
    If txtMonthly_Rate.Value > 0 Then
        If Trim(tdbWrkDays.Text) <> "" And Not IsNull(tdbWrkDays.SelectedItem) And tdbWrkDays.ApproxCount > 0 Then
            txtDaily_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text), "#,##0.0000000")
            txtHourly_Rate.Text = Format(txtMonthly_Rate.Value / CDbl(tdbWrkDays.Text) / 8, "#,##0.0000000")
        End If
    End If
End Sub

Private Sub txtPhilEr_GotFocus()
  With txtPhilEr
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPhilEr_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtPhilHAmt_GotFocus()
  With txtPhilHAmt
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPhilHAmt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtPhilHNo_GotFocus()
  With txtPhilHNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPhilHNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSection_GotFocus()
  With txtSection
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSection_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdSection_Click
End Sub

Private Sub txtSection_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSSSAmt_GotFocus()
  With txtSSSAmt
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSSSAmt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSssEc_GotFocus()
  With txtSssEc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSssEc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSssEr_GotFocus()
  With txtSssEr
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSssEr_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSSSno_GotFocus()
  With txtSSSno
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSSSno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtStreet_GotFocus()
  With txtStreet
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtStreet_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtTaxAmt_GotFocus()
  With txtTaxAmt
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtTaxAmt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtTelno_GotFocus()
  With txtTelno
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtTelno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtTinno_GotFocus()
   With txtTinno
    .SelStart = 0
    .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtTinno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtWT_GotFocus()
  With txtWT
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtWT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdWT_Click
End Sub

Private Sub txtWT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub optHDMFAuto_Click()
    If optHDMFAuto.Value = True Then
        txtHdmfAmt.Enabled = False
        txtHDMFEr.Enabled = False
    End If
End Sub

Private Sub optHDMFFixed_Click()
    If optHDMFFixed.Value = True Then
        txtHdmfAmt.Enabled = True
        txtHDMFEr.Enabled = True
    End If
End Sub

Private Sub optPhilHAuto_Click()
    If optPhilHAuto.Value = True Then
        txtPhilHAmt.Enabled = False
        txtPhilEr.Enabled = False
    End If
End Sub

Private Sub optPhilHFixed_Click()
    If optPhilHFixed.Value = True Then
        txtPhilHAmt.Enabled = True
        txtPhilEr.Enabled = True
    End If
End Sub

Private Sub optSSSAuto_Click()
    If optSSSAuto.Value = True Then
        txtSSSAmt.Enabled = False
        txtSssEr.Enabled = False
        txtSssEc.Enabled = False
    End If
End Sub

Private Sub optSSSFixed_Click()
    If optSSSFixed.Value = True Then
        txtSSSAmt.Enabled = True
        txtSssEr.Enabled = True
        txtSssEc.Enabled = True
    End If
End Sub

Private Sub optTaxAuto_Click()
    If optTaxAuto.Value = True Then
        cmdWT.Enabled = True
        txtTaxAmt.Enabled = False
    End If
End Sub

Private Sub optTaxFixed_Click()
    If optTaxFixed.Value = True Then
        cmdWT.Enabled = False
        txtTaxAmt.Enabled = True
    End If
End Sub

Private Sub Delete_Employee()
  
  Dim mDelete         As Boolean
  
  Dim rsTmp           As ADODB.Recordset
  
  mDelete = True
  
  NetOpen rsTmp, "select employeecode from loans where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from lvhdr where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from overtimehdr where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from overtimelne where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from earnings where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from deductions where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  
  NetOpen rsTmp, "select employeecode from dtr where employeecode = " & mEmployeeCode & " limit 1"
  If rsTmp.RecordCount > 0 Then
    mDelete = False
    GoTo Delete_Rec
  End If
  

Delete_Rec:
  
  If mDelete Then
    If MsgBox("Do you want to delete this record?", vbQuestion + vbYesNo) = vbYes Then
      ConMain.Execute "Delete from employee where employeecode = " & mEmployeeCode & ""
    End If
  Else
    MsgBox "You can not delete this record because it has previous payroll transacations.", vbExclamation + vbOKOnly
  End If
End Sub
