VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBrowseEmployee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Employee"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   ControlBox      =   0   'False
   Icon            =   "frmBrowseEmployee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSearch 
      BackColor       =   &H00808080&
      ForeColor       =   &H00404040&
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   -90
      Width           =   9870
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   4530
         TabIndex        =   1
         Top             =   195
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   529
         Caption         =   "frmBrowseEmployee.frx":0CCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBrowseEmployee.frx":0D36
         Key             =   "frmBrowseEmployee.frx":0D54
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
      Begin TrueOleDBList80.TDBCombo tdbSearch 
         Height          =   345
         Left            =   825
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   165
         Width           =   2730
         _ExtentX        =   4815
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
         Columns(0).DataField=   "code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "description"
         Columns(1).DataField=   "description"
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
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
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
         _PropDict       =   $"frmBrowseEmployee.frx":0D98
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H404040&"
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
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SORT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   -255
         TabIndex        =   8
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   3465
         TabIndex        =   7
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      TabIndex        =   5
      Top             =   5685
      Width           =   9840
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   75
         TabIndex        =   4
         Top             =   45
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   688
         Caption         =   "&Close"
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
         Image           =   "frmBrowseEmployee.frx":0E42
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   2115
         TabIndex        =   3
         Top             =   45
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   688
         Caption         =   "&OK"
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
         Image           =   "frmBrowseEmployee.frx":1B1C
         cBack           =   14737632
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgEmployee 
      Height          =   5160
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9102
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Employee Number"
      Columns(0).DataField=   "dummycode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Last Name"
      Columns(1).DataField=   "lastname"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "First Name"
      Columns(2).DataField=   "firstname"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Middle Name"
      Columns(3).DataField=   "middlename"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2328"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2249"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8705"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4630"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4551"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8704"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=4604"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4524"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      BorderStyle     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmBrowseEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mBrowseType          As String

Public mYear                As Integer

Dim rsEmployee              As ADODB.Recordset
Dim rsLoans                 As ADODB.Recordset
Dim rsLeaves                As ADODB.Recordset
Dim rsLeaveLimit            As ADODB.Recordset
Dim rsVoucher               As ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    'On Error Resume Next
    
    Dim CrxRep              As CRAXDRT.Report
    Dim CrxApp              As CRAXDRT.Application
    Dim crxDatabase         As CRAXDRT.Database
    Dim crxDatabaseTable    As CRAXDRT.DatabaseTable
    Dim crxDatabaseTables   As CRAXDRT.DatabaseTables
            
    If mBrowseType = "Loans" Then
    
        With frmADLoans
            
            If rsEmployee.RecordCount > 0 Then
                If Not rsEmployee.EOF Then
                
                    .mEmployeeCode = rsEmployee!employeecode
                    
                    .mBranchCode = IIf(IsNull(rsEmployee!branchcode), "Null", rsEmployee!branchcode)
                    .mDivisionCode = IIf(IsNull(rsEmployee!divisioncode), "Null", rsEmployee!divisioncode)
                    .mCostCenterCode = IIf(IsNull(rsEmployee!costcentercode), "Null", rsEmployee!costcentercode)
                    
                    NetOpen rsLoans, "select  x1.*,x2.loantypesname,(select balance from loanded where loancode = x1.loancode and fnlz = 'Y' and cancelled = 'N' order by loandedcode desc limit 1)  balance " & _
                                        "from loans x1 left outer join loantypes x2 on x1.loantypescode = x2.loantypescode where x1.employeecode = " & rsEmployee!employeecode & " order by x1.status,x1.loancode desc"
                    Set .rsLoans = rsLoans.Clone
                    Set .tdgLoan.DataSource = .rsLoans
                    
                    .txtFullname.Text = rsEmployee!lastname & ", " & rsEmployee!firstname & " " & rsEmployee!middlename
                    
                    .Get_LoanSum
                    
                    Unload Me
                    
                Else
                    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                    tdgEmployee.SetFocus
                End If
            Else
                MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                tdgEmployee.SetFocus
            End If
        End With
    
    ElseIf mBrowseType = "Leaves" Then
    
        With frmLOBLeave
            If rsEmployee.RecordCount > 0 Then
                If Not rsEmployee.EOF Then
                
                    .mEmployeeCode = rsEmployee!employeecode
                    .mLvHrParam_ID = rsEmployee!lvhrparam_id
                    .mBranchCode = IIf(IsNull(rsEmployee!branchcode), "Null", rsEmployee!branchcode)
                    .mDivisionCode = IIf(IsNull(rsEmployee!divisioncode), "Null", rsEmployee!divisioncode)
                    .mCostCenterCode = IIf(IsNull(rsEmployee!costcentercode), "Null", rsEmployee!costcentercode)

                    NetOpen rsLeaves, "select x1.*,concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname, " & _
                        "x3.costcenter,x4.division,x5.branch,x6.leavetypesname " & _
                        "from (select * from leaveapp_headers where employeecode =" & rsEmployee!employeecode & " and year(trnxdatetime) = '" & mYear & "')  x1 " & _
                        "left outer join employee x2 on x1.employeecode = x2.employeecode " & _
                        "left outer join costcenter x3 on x1.costcentercode = x3.costcentercode " & _
                        "left outer join division x4 on x1.divisioncode = x4.divisioncode " & _
                        "left outer join Branch x5 on x1.branchcode = x5.branchcode " & _
                        "left outer join leavetypes x6 on x1.leavetypescode=x6.leavetypescode " & _
                        "where x1.canceled = 'N'  " & _
                        "order by x1.leaveapp_id desc"
                   
                    Set .rsLeaves = rsLeaves.Clone
                    Set .tdgLeaves.DataSource = .rsLeaves
                    
                    If rsLeaves.RecordCount > 0 Then
                      Lock_Button "TTTT", .cmdMenu, 3
                    Else
                      Lock_Button "TFFT", .cmdMenu, 3
                    End If
                    
                    .txtFullname.Text = rsEmployee!lastname & ", " & rsEmployee!firstname & " " & rsEmployee!middlename
                    
                    .Create_TmpLeaveLimit
                    
                    NetOpen rsLeaveLimit, "select x1.leavetypescode,x1.leavetypesname,x2.lvlimit from leavetypes x1 " & _
                                          "left outer join (select * from lvlimit where employeecode = " & rsEmployee!employeecode & " and payyear = " & frmLOBLeave.txtPayYear.Text & ") x2 on x1.leavetypescode = x2.leavetypescode " & _
                                          "order by x1.leavetypesname"
                    
                    If rsLeaveLimit.RecordCount > 0 Then
                        Do While Not rsLeaveLimit.EOF
                            
                            .rsLeaveLimit.AddNew
                            .rsLeaveLimit.Fields("leavetypescode") = rsLeaveLimit!leavetypescode
                            .rsLeaveLimit.Fields("leavetypesname") = rsLeaveLimit!leavetypesname
                            .rsLeaveLimit.Fields("lvlimit") = IIf(IsNull(rsLeaveLimit!lvlimit), 0, rsLeaveLimit!lvlimit)
                            .rsLeaveLimit.Update
                            rsLeaveLimit.MoveNext
                        Loop
                    End If
                    Unload Me
                    
                Else
                    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                    tdgEmployee.SetFocus
                End If
            Else
                MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                tdgEmployee.SetFocus
            End If
        End With
    
    ElseIf mBrowseType = "Otherded" Then
        With frmUtilImportOtherded
            
            If rsEmployee.RecordCount > 0 Then
                If Not rsEmployee.EOF Then
                
                    .fg.TextMatrix(.fg.Row, 0) = rsEmployee!lastname & ", " & rsEmployee!firstname
                    .fg.Cell(flexcpBackColor, .fg.Row, 0, .fg.Row, 0) = vbWhite
                    Unload Me
                    
                Else
                    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                    tdgEmployee.SetFocus
                End If
            Else
                MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                tdgEmployee.SetFocus
            End If
        End With
        
    ElseIf mBrowseType = "Voucher" Then
    
        With frmAdVoucher
                
            If rsEmployee.RecordCount > 0 Then
                
                If Not rsEmployee.EOF Then
                    
                    .mEmployeeCode = rsEmployee!employeecode
                    
                    .mBranchCode = IIf(IsNull(rsEmployee!branchcode), "Null", rsEmployee!branchcode)
                    .mDivisionCode = IIf(IsNull(rsEmployee!divisioncode), "Null", rsEmployee!divisioncode)
                    .mCostCenterCode = IIf(IsNull(rsEmployee!costcentercode), "Null", rsEmployee!costcentercode)
                    
                    NetOpen rsVoucher, "select * from voucher x1 where x1.employeecode = " & rsEmployee!employeecode & " order by x1.vouchercode desc"
                    Set .rsVoucher = rsVoucher.Clone
                    Set .tdgLoan.DataSource = .rsVoucher
                    
                    .txtFullname.Text = rsEmployee!lastname & ", " & rsEmployee!firstname & " " & rsEmployee!middlename
                    
                    Set CrxApp = Nothing
                    Set CrxRep = Nothing
                    
                    Set CrxApp = New CRAXDRT.Application
                    Set CrxRep = New CRAXDRT.Report
                        
                    Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\rptVoucher.rpt")
                    
                    Set crxDatabase = CrxRep.Database
                    Set crxDatabaseTables = crxDatabase.Tables
    
                    For Each crxDatabaseTable In crxDatabaseTables
                        crxDatabaseTable.ConnectionProperties("connection string") = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & SQLServerName & "; DATABASE=" & SQLDatabase & "; UID=" & SQLUsername & "; PWD=" & SQLPassword & "; PORT='" & SQLPort & "'"
                    Next crxDatabaseTable
                    
                    CrxRep.ParameterFields.GetItemByName("mVoucherCode").AddCurrentValue 0
                    CrxRep.ParameterFields.GetItemByName("mVoucherType").AddCurrentValue ""
                    
                    
                    .xrpt.ReportSource = CrxRep
                    .xrpt.ViewReport
                    .xrpt.Zoom 100
                    
                    Set crxDatabase = Nothing
                    Set crxDatabaseTable = Nothing
                    Set crxDatabaseTables = Nothing
                    Set CrxApp = Nothing
    
                    Unload Me
                    
                Else
                    
                    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                    tdgEmployee.SetFocus
                    
                End If
                    
            Else
                    
                MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                tdgEmployee.SetFocus
                    
            End If
            
        End With
    
    ElseIf mBrowseType = "Employee Masterdata" Then
      
      With frmMDEmployee
          If rsEmployee.RecordCount > 0 Then
            If Not rsEmployee.EOF Then
              .ClearText
              .mEmployeeCode = rsEmployee!employeecode
              .AssignValue
              Unload Me
            Else
              MsgBox "Please select an employee.", vbExclamation + vbOKOnly
              tdgEmployee.SetFocus
            End If
          Else
            MsgBox "Please select an employee.", vbExclamation + vbOKOnly
            tdgEmployee.SetFocus
          End If
      End With
    
    ElseIf mBrowseType = "FingerPrintReg" Then
      
      With frmMDFingerPrintRegistration
          If rsEmployee.RecordCount > 0 Then
            If Not rsEmployee.EOF Then
              .Reset_Template
              .txtLastname.Text = rsEmployee!lastname & ""
              .txtFirstname.Text = rsEmployee!firstname & ""
              .txtMiddleName.Text = rsEmployee!middlename & ""
              .mEmployeeCode = rsEmployee!employeecode
              .txtPrompt.Text = "Touch the fingerprint reader."
              Unload Me
            Else
              MsgBox "Please select an employee.", vbExclamation + vbOKOnly
              tdgEmployee.SetFocus
            End If
          Else
            MsgBox "Please select an employee.", vbExclamation + vbOKOnly
            tdgEmployee.SetFocus
          End If
      End With
      
    End If
    
End Sub

Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub Form_Load()

    Dim rsTmp       As ADODB.Recordset
    
    Dim i           As Integer

    Me.MousePointer = vbHourglass
    
    CreateTmpDB rsTmp
    
    With tdgEmployee
        For i = .Columns("dummycode").ColIndex To .Columns("middlename").ColIndex
            If .Columns(i).Visible = True Then
                rsTmp.AddNew
                rsTmp.Fields("code") = .Columns(i).DataField
                rsTmp.Fields("description") = .Columns(i).Caption
                rsTmp.Update
            End If
        Next
    End With
    
    With tdbSearch
        .RowSource = rsTmp
        .ListField = "description"
        .BoundColumn = "code"
        .Columns(0).DataField = "code"
        .Columns(1).DataField = "description"
        .BoundText = "lastname"
    End With
    
    Set rsTmp = Nothing

    'If mBrowseType = "Leaves" Then
        NetOpen rsEmployee, "select employeecode,dummycode,lastname,firstname,middlename,branchcode,divisioncode,costcentercode,lvhrparam_id from employee order by lastname,firstname,middlename"
    'Else
        'NetOpen rsEmployee, "select employeecode,dummycode,lastname,firstname,middlename,branchcode,divisioncode,costcentercode from employee where isactive = 'Y' order by lastname,firstname,middlename"
    'End If
    Set tdgEmployee.DataSource = rsEmployee

    Me.MousePointer = vbDefault
    
End Sub

Private Sub tdbSearch_ItemChange()
    rsEmployee.Sort = tdbSearch.BoundText
End Sub

Private Sub tdbSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbSearch, tdbSearch.RowSource, tdbSearch.Text
        tdbSearch_ItemChange
    End If
End Sub

Private Sub tdgEmployee_DblClick()
    cmdOK_Click
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    'SendKeys "{TAB}"
    cmdOK.SetFocus
  Else
    SearchRecord KeyAscii, txtSearch, rsEmployee, txtSearch.Text, tdbSearch.BoundText
  End If
End Sub




