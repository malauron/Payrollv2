VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADOT 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   12885
   Tag             =   "Overtime Entry"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   225
      TabIndex        =   12
      Top             =   6600
      Width           =   4215
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   390
         Left            =   75
         TabIndex        =   5
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
         Image           =   "frmADOT.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   2085
         TabIndex        =   4
         Top             =   45
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   688
         Caption         =   "&Save"
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
         Image           =   "frmADOT.frx":0CDA
         cBack           =   14737632
      End
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00808080&
      Height          =   1065
      Left            =   60
      TabIndex        =   8
      Top             =   795
      Width           =   12825
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1785
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   195
         Width           =   3900
         _ExtentX        =   6879
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
         Columns(0).DataField=   "percode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   "description"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "from"
         Columns(2).DataField=   "wrkdatefrom"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "To"
         Columns(3).DataField=   "wrkdateto"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "payyear"
         Columns(4).DataField=   "payyear"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "paymonth"
         Columns(5).DataField=   "paymonth"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "payfreqcode"
         Columns(6).DataField=   "payfreqcode"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1958"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1879"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2328"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2249"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
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
         _PropDict       =   $"frmADOT.frx":19B4
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   6900
         TabIndex        =   2
         Top             =   630
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   529
         Caption         =   "frmADOT.frx":1A5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADOT.frx":1ACA
         Key             =   "frmADOT.frx":1AE8
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
      Begin TrueOleDBList80.TDBCombo tdbSort 
         Height          =   345
         Left            =   1785
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   600
         Width           =   3900
         _ExtentX        =   6879
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
         _PropDict       =   $"frmADOT.frx":1B2C
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   5310
         TabIndex        =   11
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "PAYROLL PERIOD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "SORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   645
         Width           =   465
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
      ScaleWidth      =   12885
      TabIndex        =   6
      Top             =   0
      Width           =   12885
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime"
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
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   5445
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgOvertime 
      Height          =   4005
      Left            =   60
      TabIndex        =   3
      Top             =   1875
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   7064
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "empno"
      Columns(0).DataField=   "employeecode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Employee"
      Columns(1).DataField=   "employeename"
      Columns(1).DropDown=   "tddEmployee"
      Columns(1).DropDown.vt=   8
      Columns(1).ExternalEditor=   "txtEmployee"
      Columns(1).ExternalEditor.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Date"
      Columns(2).DataField=   "wrkdate"
      Columns(2).NumberFormat=   "MM/DD/YYYY"
      Columns(2).ExternalEditor=   "txtWrkdate"
      Columns(2).ExternalEditor.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "In"
      Columns(3).DataField=   "otstart"
      Columns(3).NumberFormat=   "hh:nn"
      Columns(3).ExternalEditor=   "txtTime"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Out"
      Columns(4).DataField=   "otend"
      Columns(4).NumberFormat=   "hh:nn"
      Columns(4).ExternalEditor=   "txtTime"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Total Hrs"
      Columns(5).DataField=   "othrs"
      Columns(5).NumberFormat=   "#,##0.00"
      Columns(5).ExternalEditor=   "txt2"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Remarks"
      Columns(6).DataField=   "remarks"
      Columns(6).ExternalEditor=   "txtRemarks"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=7752"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7673"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1244"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1164"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1191"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1111"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1931"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1852"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=8096"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerStyle=0"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=8043"
      Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(45)=   "Column(6)._HeadDivider=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Width=79"
      Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
      Appearance      =   2
      BorderStyle     =   0
      DefColWidth     =   0
      HeadLines       =   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=0"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin TDBTime6Ctl.TDBTime txtTime 
      Height          =   285
      Left            =   9090
      TabIndex        =   13
      Top             =   8085
      Visible         =   0   'False
      Width           =   1650
      _Version        =   65536
      _ExtentX        =   2910
      _ExtentY        =   503
      Caption         =   "frmADOT.frx":1BD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmADOT.frx":1C42
      Spin            =   "frmADOT.frx":1C92
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "14:25"
      ValidateMode    =   0
      ValueVT         =   1935998983
      Value           =   0.600914351851852
   End
   Begin TDBDate6Ctl.TDBDate txtWrkdate 
      Height          =   285
      Left            =   7350
      TabIndex        =   14
      Top             =   7845
      Visible         =   0   'False
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   503
      Calendar        =   "frmADOT.frx":1CBA
      Caption         =   "frmADOT.frx":1DD2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":1E3E
      Keys            =   "frmADOT.frx":1E5C
      Spin            =   "frmADOT.frx":1EBA
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
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
      Text            =   "04/02/2008"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39540
      CenturyMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtStatus 
      Height          =   285
      Left            =   8970
      TabIndex        =   15
      Top             =   6990
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADOT.frx":1EE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":1F4E
      Key             =   "frmADOT.frx":1F6C
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
      BorderStyle     =   0
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
   Begin TDBText6Ctl.TDBText txtSignatory 
      Height          =   285
      Left            =   5745
      TabIndex        =   16
      Top             =   7395
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADOT.frx":1FB0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":201C
      Key             =   "frmADOT.frx":203A
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
      BorderStyle     =   0
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
   Begin TrueOleDBGrid80.TDBDropDown tddStatus 
      Height          =   1365
      Left            =   570
      TabIndex        =   17
      Top             =   7050
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Name"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   2
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   0
      RowDividerStyle =   0
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   -1  'True
      ListField       =   "Branch"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   -1  'True
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFFF0EA&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFFFF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown tddSignatory 
      Height          =   1365
      Left            =   6840
      TabIndex        =   18
      Top             =   6345
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Name"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   2
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   0
      RowDividerStyle =   0
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   -1  'True
      ListField       =   "Branch"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   -1  'True
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFFF0EA&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFFFF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown tddEmployee 
      Height          =   4260
      Left            =   3585
      TabIndex        =   19
      Top             =   3810
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   7514
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Name"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   2
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   0
      RowDividerStyle =   0
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   -1  'True
      ListField       =   "Branch"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   -1  'True
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFFF0EA&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFFFF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText txtRemarks 
      Height          =   285
      Left            =   4335
      TabIndex        =   20
      Top             =   8790
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADOT.frx":207E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":20EA
      Key             =   "frmADOT.frx":2108
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
      BorderStyle     =   0
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
   Begin TDBText6Ctl.TDBText txtEmployee 
      Height          =   285
      Left            =   2460
      TabIndex        =   21
      Top             =   8430
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADOT.frx":214C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":21B8
      Key             =   "frmADOT.frx":21D6
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
      BorderStyle     =   0
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
   Begin TDBNumber6Ctl.TDBNumber txt2 
      Height          =   300
      Left            =   0
      TabIndex        =   22
      Top             =   8430
      Visible         =   0   'False
      Width           =   1470
      _Version        =   65536
      _ExtentX        =   2593
      _ExtentY        =   529
      Calculator      =   "frmADOT.frx":221A
      Caption         =   "frmADOT.frx":223A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADOT.frx":22A0
      Keys            =   "frmADOT.frx":22BE
      Spin            =   "frmADOT.frx":2308
      AlignHorizontal =   1
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,##0.00"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   186843137
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
End
Attribute VB_Name = "frmADOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public rsOvertime           As ADODB.Recordset
Dim rsTempOTEntry           As ADODB.Recordset
Dim mPerCode                As String

Private Sub cmdmenu_Click(Index As Integer)

  Select Case Index

    Case 0:

          If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then
                frmADOvertime2.mAdd = True
                frmADOvertime2.Show vbModal
          Else
                MsgBox "Please choose a payrol period.", vbExclamation + vbOKOnly
          End If

    Case 1:

          If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then
            frmADOvertime2.mAdd = False
            frmADOvertime2.Show vbModal
          Else
            MsgBox "Please choose a payroll period.", vbExclamation + vbOKOnly
          End If

    Case 2:
    Case 3:
    Case 4: Unload Me

  End Select

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsChk               As ADODB.Recordset
    Dim rsDateTime          As ADODB.Recordset

    Dim mDate               As String
    Dim mWorkDate           As String
    Dim mEntrdBU            As String
    Dim mIn                 As String
    Dim mOut                As String
    

    If rsTempOTEntry Is Nothing Then
      MsgBox "You must provide at least one (1) overtime entry.", vbExclamation + vbOKOnly
      tdgOvertime.SetFocus
      Exit Sub
    End If

    If rsTempOTEntry.RecordCount <= 0 Then
      MsgBox "You must provide at least one (1) overtime entry.", vbExclamation + vbOKOnly
      tdgOvertime.SetFocus
      Exit Sub
    End If

    With rsTempOTEntry
      .MoveFirst
      Do While Not .EOF
      
        If Not IsDate(!wrkdate) Then
'          MsgBox "Please check the data for invalid date format.", vbExclamation + vbOKOnly
'          tdgOvertime.SetFocus
'          Exit Sub
            mWorkDate = "Null"
        Else
            mWorkDate = "'" & Format(!wrkdate, "YYYY-MM-DD") & "'"
        End If

        If Not IsDate(!otstart) Then
'          MsgBox "Please check the data for invalid time format.", vbExclamation + vbOKOnly
'          tdgOvertime.SetFocus
'          Exit Sub
            mIn = "Null"
        Else
            mIn = "'" & Format(!otstart, "hh:nn") & "'"
        End If

        If Not IsDate(!otend) Then
'          MsgBox "Please check the data for invalid time format.", vbExclamation + vbOKOnly
'          tdgOvertime.SetFocus
'          Exit Sub
            mOut = "Null"
        Else
            mOut = "'" & Format(!otend, "hh:nn") & "'"
        End If
        
        If Not IsNumeric(!othrs) Then
            MsgBox "Number of OT hrs must be greater than zero.", vbExclamation + vbOKOnly
            tdgOvertime.SetFocus
            Exit Sub
        End If

        .MoveNext
        
      Loop
    End With

    If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If

    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans

    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"
    
    ConMain.Execute "delete from overtimelne where percode = " & mPerCode & ""
    
    With rsTempOTEntry

        .MoveFirst

        Do While Not .EOF
        
            If IsNull(!tdatetime) Then
                mDate = "Null"
            Else
                mDate = "'" & Format(!tdatetime, "YYYY-MM-DD hh:nn:ss") & "'"
            End If
            
            If Not IsDate(!wrkdate) Then
                mWorkDate = "Null"
            Else
                mWorkDate = "'" & Format(!wrkdate, "YYYY-MM-DD") & "'"
            End If
    
            If Not IsDate(!otstart) Then
                mIn = "Null"
            Else
                mIn = "'" & Format(!otstart, "hh:nn") & "'"
            End If
    
            If Not IsDate(!otend) Then
                mOut = "Null"
            Else
                mOut = "'" & Format(!otend, "hh:nn") & "'"
            End If
    
            If Trim(!enteredbyuser) = "" Then
                mEntrdBU = "N"
            Else
                mEntrdBU = !enteredbyuser
            End If
            
            ConMain.Execute "insert into overtimelne(employeecode,percode,wrkdate,approvby, " & _
                              "remarks,otstart,otend,othrs, " & _
                              "status,tdatetime,enteredbyuser,fnlz) values " & _
                              "('" & !EmployeeCode & "','" & mPerCode & "'," & mWorkDate & ",'" & !approvby & "', " & _
                              "'" & !remarks & "'," & mIn & "," & mOut & "," & !othrs & ", " & _
                              "'Approved'," & mDate & ",'" & mEntrdBU & "','N')"

        .MoveNext

        Loop

    End With

    ConMain.CommitTrans

    MsgBox "Data was successfully saved.", vbInformation + vbOKOnly
    tdgOvertime.SetFocus
    
End Sub

Private Sub Form_Load()

    Dim I             As Integer
    Dim rsTMP         As ADODB.Recordset

    Add_MDIButton Me.Name, Me.Tag

    bind_tdb ConMain, tdbPayrollPeriod, "select percode, description, wrkdatefrom, wrkdateto,payyear,paymonth,payfreqcode from payrollperiod order by percode desc", "description", "percode"

    CreateTmpDB rsTMP

    With rsTMP
      For I = 1 To 2
        .AddNew
        Select Case I
            Case 2: .Fields("code") = "employeename"
                    .Fields("description") = "Fullname"
            Case 1: .Fields("code") = "wrkdate"
                    .Fields("description") = "Date"
        End Select
        .Update
      Next
    End With

    With tdbSort
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTMP
    End With

    CreateTmpDB rsTMP

    With rsTMP
      For I = 1 To 3
          .AddNew
          Select Case I
              Case 1: .Fields("code") = "Approved"
                      .Fields("description") = "Approved"
              Case 2: .Fields("code") = "Cancelled"
                      .Fields("description") = "Cancelled"
              Case 3: .Fields("code") = "Pending"
                      .Fields("description") = "Pending"
          End Select
          .Update
      Next
    End With

    With tddStatus
      .DataSource = rsTMP
      .ListField = "description"
    End With

    Set rsTMP = Nothing

    Bind_tdd ConMain, tddEmployee, "select employeecode code,concat(lastname,', ',firstname,' ',middlename) description from employee " & _
                              "where payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' " & _
                              "order by lastname,firstname,middlename", "description"

    Bind_tdd ConMain, tddSignatory, "select fullname code,fullname description from signatory order by fullname", "description"

    Set rsTMP = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

  On Error Resume Next


    With fra1
      .Top = pic1.Top + pic1.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With

    With tdgOvertime
      .Top = fra1.Top + fra1.Height
      .Left = 0
      .Height = Me.ScaleHeight - (.Top + fraButtons.Height)
      .Width = Me.ScaleWidth
    End With

    With fraButtons
        .Top = tdgOvertime.Top + tdgOvertime.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With

End Sub

Private Sub tdbPayrollPeriod_KeyPress(Keyascii As Integer)
  If Keyascii = 13 Then
    If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then
        Load_OT
    End If
    SendKeys "{TAB}"
  Else
    SearchList Keyascii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
  End If
End Sub

Private Sub tdbPayrollPeriod_ItemChange()

  Set rsTempOTEntry = Nothing
  Set tdgOvertime.DataSource = Nothing

End Sub

Private Sub tdbSort_KeyPress(Keyascii As Integer)
  If Keyascii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList Keyascii, tdbSort, tdbSort.RowSource, tdbSort.Text
    tdbSort_ItemChange
  End If
End Sub

Private Sub tdbSort_ItemChange()
    With rsTempOTEntry
        
        If Not rsTempOTEntry Is Nothing Then
            .Sort = tdbSort.BoundText
        End If
        
    End With
End Sub

Private Sub tddEmployee_RowChange()
    With tddEmployee
        tdgOvertime.Columns("employeecode").Text = .Columns("code").Text
        'tdgOvertime.Columns("employeename").Text = .Columns("description").Text
    End With
End Sub

Private Sub tddEmployee_DropDownOpen()

    With tddEmployee
        Bind_tdd ConMain, tddEmployee, "select employeecode code,concat(lastname,', ',firstname,' ',middlename) description from employee " & _
                              "where payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' " & _
                              "order by lastname,firstname,middlename", "description"
        If Trim(txtEmployee.Text) = "" Then
            txtEmployee.Text = .Columns("description").Text
        End If
        .Width = tdgOvertime.Columns("employeename").Width
        .Height = tdgOvertime.Height - 800
    End With

End Sub

Private Sub tdgOvertime_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    With tdgOvertime
        If ColIndex = .Columns("otstart").ColIndex Or ColIndex = .Columns("otend").ColIndex Then
            If IsDate(.Columns("wrkdate").Text) Then
                .Columns("othrs").Text = TtlHrs(.Columns("wrkdate").Text & " " & .Columns("otstart").Text, .Columns("wrkdate").Text & " " & .Columns("otend").Text)
            End If
'        ElseIf ColIndex = .Columns("wrkdate").ColIndex Then
'            If IsDate(.Columns("wrkdate").Text) Then
'                If Trim(.Columns("status").Text) = "" Then
'                    .Columns("status").Text = "Approved"
'                End If
'            End If
        End If
    End With

End Sub

Private Sub tdgOvertime_KeyDown(KeyCode As Integer, Shift As Integer)
    
    With tdgOvertime
            If txtEmployee.Visible = False And txtWrkdate.Visible = False And txtTime.Visible = False Then
                If .ApproxCount > 0 Then
                    If Not .EOF Then
                        If KeyCode = 46 Then
                            If MsgBox("Do you want to delete this entry.", vbQuestion + vbYesNo) = vbYes Then
                              .Delete
                              .Refresh
                            End If
                            .SetFocus
                        End If
                    End If
                End If
            End If
    End With

End Sub

Private Sub tdgOvertime_KeyPress(Keyascii As Integer)

    On Error GoTo ErrHndlr
    With tdgOvertime
        If Keyascii = 13 Then
            If .Col - 1 = .Columns("remarks").ColIndex Then
                If .Row < .ApproxCount - 1 Then
                    .Row = .Row + 1
                    .Col = .Columns("employeename").ColIndex
                ElseIf .Row = .ApproxCount - 1 Then
                    SendKeys "{DOWN}"
                    .Col = .Columns("employeename").ColIndex
                ElseIf .Row > .ApproxCount - 1 Then
                    .Col = .Columns("remarks").ColIndex
                    SendKeys "{TAB}"
                End If
            End If
        End If
    End With
ErrHndlr:
End Sub

Private Sub txt2_LostFocus()
    
    On Error Resume Next
    
    tdgOvertime.SetFocus
    
End Sub

Private Sub txtEmployee_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        tdgOvertime.SetFocus
    Else
        SearchRecord Keyascii, txtEmployee, tddEmployee.DataSource, txtEmployee.Text, "description"
        tddEmployee_RowChange
    End If
End Sub

Private Sub txtEmployee_LostFocus()
    On Error Resume Next
    tdgOvertime.SetFocus
End Sub

Private Sub txtRemarks_LostFocus()
    tdgOvertime.SetFocus
End Sub

Private Sub txtSearch_KeyPress(Keyascii As Integer)
  If Keyascii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchRecord Keyascii, txtSearch, rsTempOTEntry, txtSearch.Text, tdbSort.BoundText
  End If
End Sub

Private Sub Create_TmpOTEntry()

'    If Not rsTempOTEntry Is Nothing Then
'        If rsTempOTEntry.RecordCount > 0 Then
'            rsTempOTEntry.MoveFirst
'            Do While Not rsTempOTEntry.EOF
'                rsTempOTEntry.Delete
'                rsTempOTEntry.Update
'                If rsTempOTEntry.RecordCount > 0 Then
'                    rsTempOTEntry.MoveNext
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'    End If

    Set rsTempOTEntry = Nothing
    Set rsTempOTEntry = New ADODB.Recordset

    With rsTempOTEntry

        .Fields.Append "otlneno", adInteger, , adFldIsNullable
        .Fields.Append "otcode", adInteger, , adFldIsNullable
        .Fields.Append "percode", adInteger, , adFldIsNullable
        .Fields.Append "employeecode", adInteger, , adFldIsNullable
        .Fields.Append "dummycode", adVarChar, 10
        .Fields.Append "employeename", adVarChar, 300
        .Fields.Append "wrkdate", adDate, , adFldIsNullable
        .Fields.Append "approvby", adVarChar, 70
        .Fields.Append "remarks", adVarChar, 100
        .Fields.Append "otstart", adVarChar, 11, adFldIsNullable
        .Fields.Append "otend", adVarChar, 11, adFldIsNullable
        .Fields.Append "othrs", adDouble, 18, adFldIsNullable
        .Fields.Append "status", adVarChar, 20
        .Fields.Append "enteredbyuser", adVarChar, 1
        .Fields.Append "tdatetime", adDate, , adFldIsNullable
        .Open

    End With

    Set tdgOvertime.DataSource = rsTempOTEntry

End Sub

Private Sub Load_OT()

    Dim rsOTEntry     As ADODB.Recordset

    Create_TmpOTEntry
    
    mPerCode = tdbPayrollPeriod.Columns("percode").Text

    NetOpen rsOTEntry, "select x1.*,x2.dummycode,concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) employeename from overtimelne x1 " & _
                       "left outer join employee x2 on x1.employeecode = x2.employeecode where x1.fnlz = 'N' and x1.percode = " & tdbPayrollPeriod.Columns("percode").Text & ""

    With rsOTEntry
        If .RecordCount > 0 Then

          .MoveFirst

          Do While Not .EOF

            rsTempOTEntry.AddNew
            rsTempOTEntry.Fields("otlneno") = !otlneno
            rsTempOTEntry.Fields("otcode") = !otcode
            rsTempOTEntry.Fields("percode") = !percode
            rsTempOTEntry.Fields("employeecode") = !EmployeeCode
            rsTempOTEntry.Fields("dummycode") = !dummycode
            rsTempOTEntry.Fields("employeename") = !employeename
            rsTempOTEntry.Fields("wrkdate") = !wrkdate
            rsTempOTEntry.Fields("approvby") = !approvby
            rsTempOTEntry.Fields("remarks") = !remarks
            rsTempOTEntry.Fields("otstart") = !otstart
            rsTempOTEntry.Fields("otend") = !otend
            rsTempOTEntry.Fields("othrs") = !othrs
            rsTempOTEntry.Fields("status") = !Status
            rsTempOTEntry.Fields("enteredbyuser") = !enteredbyuser
            rsTempOTEntry.Fields("tdatetime") = IIf(IsNull(!tdatetime), Null, Format(!tdatetime, "YYYY-MM-DD"))
            rsTempOTEntry.Update

            .MoveNext

          Loop

        End If

    End With

End Sub

Private Sub tddStatus_RowChange()
    With tddStatus
        txtStatus.Text = .Columns("description").Text
    End With
End Sub

Private Sub tddStatus_DropDownOpen()
    With tddStatus
        .Width = tdgOvertime.Columns("status").Width
    End With
End Sub

Private Sub txtStatus_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        tdgOvertime.SetFocus
    Else
        SearchRecord Keyascii, txtStatus, tddStatus.DataSource, txtStatus.Text, "description"
        tddStatus_RowChange
    End If
End Sub

Private Sub txtStatus_LostFocus()
    tdgOvertime.SetFocus
End Sub

Private Function TtlHrs(ByRef objTin As String, ByRef objTout As String) As Double

  Dim mTime         As Double

  If IsDate(objTin) And IsDate(objTout) Then
    If CDate(objTin) > CDate(objTout) Then
      mTime = 24 + Format(Round(DateDiff("N", objTin, "12:00 am") / 60, 2), "#,##0.00")
      TtlHrs = Format(Round(DateDiff("N", "12:00 am", objTout) / 60, 2), "#,##0.00") + mTime
    Else
      TtlHrs = Format(Round(DateDiff("N", objTin, objTout) / 60, 2), "#,##0.00")
      If TtlHrs = 0 Then TtlHrs = 24
    End If
  Else
    TtlHrs = 0
  End If

End Function

Private Sub txtTime_LostFocus()
    tdgOvertime.SetFocus
End Sub

Private Sub txtWrkdate_LostFocus()
    tdgOvertime.SetFocus
End Sub

