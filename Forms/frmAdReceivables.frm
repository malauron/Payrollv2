VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAdReceivables 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00808080&
      Height          =   1020
      Left            =   0
      TabIndex        =   18
      Top             =   735
      Width           =   12705
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   8160
         TabIndex        =   2
         Top             =   630
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   529
         Caption         =   "frmAdReceivables.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdReceivables.frx":006C
         Key             =   "frmAdReceivables.frx":008A
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
         Left            =   1635
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   615
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
         _PropDict       =   $"frmAdReceivables.frx":00CE
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
      Begin TrueOleDBList80.TDBCombo tdbOtherDeductions 
         Height          =   345
         Left            =   1635
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   210
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
         Columns(0).Caption=   "Employee Code"
         Columns(0).DataField=   "employeecode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Employee Name"
         Columns(1).DataField=   "employeename"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=8361"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8281"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
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
         _PropDict       =   $"frmAdReceivables.frx":0178
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H0&"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
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
         Left            =   1110
         TabIndex        =   21
         Top             =   645
         Width           =   465
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
         Left            =   6420
         TabIndex        =   20
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIVABLE TYPE"
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
         Left            =   -60
         TabIndex        =   19
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   15
      TabIndex        =   10
      Top             =   5820
      Width           =   12675
      Begin TDBNumber6Ctl.TDBNumber txtTotal 
         Height          =   315
         Left            =   4935
         TabIndex        =   11
         Top             =   60
         Width           =   2205
         _Version        =   65536
         _ExtentX        =   3889
         _ExtentY        =   556
         Calculator      =   "frmAdReceivables.frx":0222
         Caption         =   "frmAdReceivables.frx":0242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdReceivables.frx":02AE
         Keys            =   "frmAdReceivables.frx":02CC
         Spin            =   "frmAdReceivables.frx":0316
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,###,##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   4210752
         Format          =   "#,###,###,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBNumber6Ctl.TDBNumber txtNoOfRecords 
         Height          =   315
         Left            =   1305
         TabIndex        =   12
         Top             =   60
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
         _ExtentY        =   556
         Calculator      =   "frmAdReceivables.frx":033E
         Caption         =   "frmAdReceivables.frx":035E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdReceivables.frx":03CA
         Keys            =   "frmAdReceivables.frx":03E8
         Spin            =   "frmAdReceivables.frx":0432
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   1
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   4210752
         Format          =   "#,###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   -999999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   15
         TabIndex        =   13
         Top             =   345
         Width           =   7395
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   150
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
            Image           =   "frmAdReceivables.frx":045A
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   1
            Left            =   1515
            TabIndex        =   4
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&EDIT"
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
            Image           =   "frmAdReceivables.frx":2134
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   2
            Left            =   2970
            TabIndex        =   14
            Top             =   765
            Visible         =   0   'False
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
            Image           =   "frmAdReceivables.frx":3E0E
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   3
            Left            =   2970
            TabIndex        =   5
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "CANCE&L"
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
            Image           =   "frmAdReceivables.frx":5AE8
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   4
            Left            =   5865
            TabIndex        =   7
            Top             =   150
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
            Image           =   "frmAdReceivables.frx":67C2
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   5
            Left            =   4425
            TabIndex        =   6
            Top             =   150
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&IMPORT"
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
            Image           =   "frmAdReceivables.frx":709C
            cBack           =   14737632
         End
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   3720
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F6F8F8&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Records"
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
         Left            =   60
         TabIndex        =   15
         Top             =   120
         Width           =   1215
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
      ScaleWidth      =   12705
      TabIndex        =   8
      Top             =   0
      Width           =   12705
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Receivables"
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
         TabIndex        =   9
         Top             =   225
         Width           =   5445
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgReceivables 
      Height          =   4005
      Left            =   15
      TabIndex        =   17
      Top             =   1770
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   7064
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Receivable Number"
      Columns(0).DataField=   "receivablecode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Employee ID"
      Columns(1).DataField=   "employeeid"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Employee Name"
      Columns(2).DataField=   "employeename"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Amount"
      Columns(3).DataField=   "amount"
      Columns(3).NumberFormat=   "#,##0.00"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Remarks"
      Columns(4).DataField=   "remarks"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1931"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=9684"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9604"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2884"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerStyle=0"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2831"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(4)._HeadDivider=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
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
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmAdReceivables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mDeductionType                     As Integer

Public rsReceivables                      As ADODB.Recordset

Private Sub cmdmenu_Click(Index As Integer)
    
    Dim mRow                              As Long
    
    Dim mTotal                            As Double
    
    Select Case Index
            
      Case 0:
            
            If mDeductionType <> 0 Then
              frmAdReceivables2.mAdd = True
              frmAdReceivables2.Show vbModal
            Else
              MsgBox "Receivable type is blank.", vbExclamation + vbOKOnly
            End If
              
      Case 1:
            
            If mDeductionType <> 0 Then
              frmAdReceivables2.mAdd = False
              frmAdReceivables2.Show vbModal
            Else
              MsgBox "Receivable type is blank.", vbExclamation + vbOKOnly
            End If
              
      Case 2:
            
      Case 3:
            
            If rsReceivables.RecordCount <= 0 Then Exit Sub
            
            If MsgBox("Do you want to delete this receivable entry?", vbInformation + vbYesNo) = vbYes Then
                
                ConMain.Execute "delete from receivables where vouchercode = " & rsReceivables!vouchercode & ""
                
                If rsReceivables.AbsolutePosition = rsReceivables.RecordCount Then
                    mRow = rsReceivables.AbsolutePosition - 1
                Else
                    mRow = rsReceivables.AbsolutePosition
                End If
                
                rsReceivables.Requery
                
                If rsReceivables.RecordCount > 0 Then
                
                    txtNoOfRecords.Text = Format(rsReceivables.RecordCount, "#,##0")
                    
                    Lock_Button "TTFTTT", cmdMenu, 5
                    rsReceivables.MoveFirst
                    Do While Not rsReceivables.EOF
                        mTotal = mTotal + rsReceivables!amount
                        rsReceivables.MoveNext
                    Loop
                    
                    rsReceivables.MoveFirst
                    'lblTotal.Caption = "Total: " & Format(mTotal, "#,##0.00")
                    txtTotal.Text = Format(mTotal, "#,##0.00")
                    
                Else
                    
                    Lock_Button "TFFFTT", cmdMenu, 5
                    'lblTotal.Caption = "Total: N/A"
                    txtNoOfRecords.Text = "0"
                    txtTotal.Text = "0.00"
                    
                End If
                
                If mRow > 0 Then
                    rsReceivables.AbsolutePosition = mRow
                End If
                
                tdgReceivables.SetFocus
                
            End If
            
      Case 4: Unload Me
            
      Case 5:
            
            If mDeductionType = 0 Then
                MsgBox "Please select a voucher type.", vbExclamation + vbOKOnly
                tdbOtherDeductions.SetFocus
                Exit Sub
            End If
            
            With frmUtilImportOtherded
                .mImportType = "receivables"
                .mDeductionType = tdbOtherDeductions.BoundText
                .Show vbModal
            End With
            
    End Select
    
End Sub
    
Private Sub Form_Load()
    
    Dim I               As Integer
    
    Dim rsTMP           As ADODB.Recordset
    Dim rsTmp2          As ADODB.Recordset
    
    Add_MDIButton Me.Name, Me.Tag
      
    CreateTmpDB rsTMP
    
    With rsTMP
      For I = 1 To 4
        .AddNew
        Select Case I
            Case 1: .Fields("code") = "receivablecode"
                    .Fields("description") = "Receivable Number"
            Case 2: .Fields("code") = "referenceno"
                    .Fields("description") = "Reference Number"
            Case 3: .Fields("code") = "employeeid"
                    .Fields("description") = "Employee ID"
            Case 4: .Fields("code") = "employeename"
                    .Fields("description") = "Fullname"
        End Select
        .Update
      Next
    End With
    
    With tdbSort
      .BoundColumn = "code"
      .ListField = "description"
      .Columns(0).DataField = "code"
      .Columns(1).DataField = "description"
      .RowSource = rsTMP
      .BoundText = "employeename"
    End With
    
    Set rsTMP = Nothing
    
    Lock_Button "FFFFTF", cmdMenu, 5
    
End Sub
    
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Remove_MDIButton Me.Name
    
    mDeductionType = 0
    
    Set rsReceivables = Nothing
    
End Sub
    
Private Sub Form_Resize()
    
    On Error Resume Next
        
    With fra1
      .Top = pic1.Top + pic1.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With tdgReceivables
      .Top = fra1.Top + fra1.Height
      .Left = 0
      .Height = Me.ScaleHeight - (.Top + fraButtons.Height)
      .Width = Me.ScaleWidth
    End With
    
    With fraButtons
      .Top = tdgReceivables.Top + tdgReceivables.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
End Sub
    
Private Sub Cancel_Clicked()
    
  If rsReceivables.RecordCount > 0 Then
    Lock_Button "TTFFTT", cmdMenu, 5
  Else
    Lock_Button "TFFFTT", cmdMenu, 5
  End If
  
End Sub

Private Sub tdbOtherDeductions_GotFocus()

    If Trim(tdbOtherDeductions.Text) <> "" And Not IsNull(tdbOtherDeductions.SelectedItem) And tdbOtherDeductions.ApproxCount > 0 Then
        tdbOtherDeductions.Tag = tdbOtherDeductions.BoundText
    Else
        tdbOtherDeductions.Tag = ""
    End If
    
    bind_tdb ConMain, tdbOtherDeductions, "select otherDeductionscode,otherDeductionsname from otherDeductions where isactive = 'Y' order by otherDeductionsname", "otherDeductionsname", "otherDeductionscode"
    
    If mDeductionType <> 0 Then
        tdbOtherDeductions.BoundText = mDeductionType
    Else
        tdbOtherDeductions.BoundText = ""
    End If
    
End Sub
    
Private Sub tdbOtherDeductions_KeyPress(Keyascii As Integer)
    
    If Keyascii = 13 Then
      
        SendKeys "{TAB}"
        
        Get_List
          
    Else
    
        SearchList Keyascii, tdbOtherDeductions, tdbOtherDeductions.RowSource, tdbOtherDeductions.Text
        
    End If
    
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
    
    On Error Resume Next
    
    rsReceivables.Sort = tdbSort.BoundText
    
End Sub

Private Sub tdgreceivables_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmdmenu_Click 3
    End If
End Sub

Private Sub txtSearch_GotFocus()
    
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtSearch_KeyPress(Keyascii As Integer)
    
    If Keyascii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchRecord Keyascii, txtSearch, rsReceivables, txtSearch.Text, tdbSort.BoundText
    End If
    
End Sub

Public Sub Get_List()

    Dim mTotal              As Double
    
    If Trim(tdbOtherDeductions.Text) <> "" And Not IsNull(tdbOtherDeductions.SelectedItem) And tdbOtherDeductions.ApproxCount > 0 Then
        mDeductionType = tdbOtherDeductions.BoundText
    Else
        mDeductionType = 0
    End If
    
    NetOpen rsReceivables, "select x1.*,x2.dummycode employeeid,concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) employeename from receivables x1 " & _
               "left outer join employee x2 on x1.employeecode = x2.employeecode " & _
               "where x1.deductioncode = '" & mDeductionType & "' order by " & tdbSort.BoundText & ""
    
    Set tdgReceivables.DataSource = rsReceivables
    
    If mDeductionType <> 0 Then
     
        If rsReceivables.RecordCount > 0 Then
        
            txtNoOfRecords.Text = Format(rsReceivables.RecordCount, "#,##0")
            
            Lock_Button "TTFTTT", cmdMenu, 5
            
            rsReceivables.MoveFirst
            Do While Not rsReceivables.EOF
                mTotal = mTotal + rsReceivables!amount
                rsReceivables.MoveNext
            Loop
            
            rsReceivables.MoveFirst
            
            'lblTotal.Caption = "Total: " & Format(mTotal, "#,##0.00")
            
            txtTotal.Text = Format(mTotal, "#,##0.00")
            
        Else
            
            Lock_Button "TFFFTT", cmdMenu, 5
            
            'lblTotal.Caption = "Total: N/A"
            
            txtNoOfRecords.Text = "0"
            txtTotal.Text = "0.00"
            
        End If
        
    Else
        
        Lock_Button "FFFFTF", cmdMenu, 5
        'lblTotal.Caption = "Total: N/A"
        txtNoOfRecords.Text = "0"
        txtTotal.Text = "0.00"
        
    End If
        
End Sub
