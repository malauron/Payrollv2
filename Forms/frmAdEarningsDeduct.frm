VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAdEarningsDeduct 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   13305
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      Height          =   1800
      Left            =   -510
      TabIndex        =   6
      Top             =   150
      Width           =   14385
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   1395
         Width           =   3900
         _Version        =   65536
         _ExtentX        =   6879
         _ExtentY        =   529
         Caption         =   "frmAdEarningsDeduct.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEarningsDeduct.frx":006C
         Key             =   "frmAdEarningsDeduct.frx":008A
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
      Begin TrueOleDBList80.TDBCombo tdbSection 
         Height          =   345
         Left            =   7185
         TabIndex        =   4
         Tag             =   "Municipal"
         Top             =   990
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
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   ""
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
         _PropDict       =   $"frmAdEarningsDeduct.frx":00CE
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
      Begin TrueOleDBList80.TDBCombo tdbBranch 
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   585
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
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   ""
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
         _PropDict       =   $"frmAdEarningsDeduct.frx":0178
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
      Begin TrueOleDBList80.TDBCombo tdbDivision 
         Height          =   345
         Left            =   1800
         TabIndex        =   2
         Tag             =   "Municipal"
         Top             =   990
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
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   ""
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
         _PropDict       =   $"frmAdEarningsDeduct.frx":0222
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
      Begin TrueOleDBList80.TDBCombo tdbCostCenter 
         Height          =   345
         Left            =   7185
         TabIndex        =   3
         Tag             =   "Municipal"
         Top             =   585
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
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   ""
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
         _PropDict       =   $"frmAdEarningsDeduct.frx":02CC
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
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   330
         Left            =   11520
         TabIndex        =   13
         Top             =   960
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         Caption         =   "&Generate"
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
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   330
         Left            =   13335
         TabIndex        =   14
         Top             =   960
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
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
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Tag             =   "Municipal"
         Top             =   180
         Width           =   5310
         _ExtentX        =   9366
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
         Columns(2).NumberFormat=   "mm/dd/yyyy"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "To"
         Columns(3).DataField=   "wrkdateto"
         Columns(3).NumberFormat=   "mm/dd/yyyy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "payfreqcode"
         Columns(4).DataField=   "payfreqcode"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "payyear"
         Columns(5).DataField=   "payyear"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "paymonth"
         Columns(6).DataField=   "paymonth"
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
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3016"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2937"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1693"
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
         HeadLines       =   0
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
         _PropDict       =   $"frmAdEarningsDeduct.frx":0376
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF6F8F8&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "Cost Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   10
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   9
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "Payroll Period"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F6F8F8&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   7
         Top             =   1425
         Width           =   1455
      End
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   7875
      TabIndex        =   5
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
   Begin VSFlex8Ctl.VSFlexGrid fgEarningsDeduct 
      Height          =   3900
      Left            =   2220
      TabIndex        =   16
      Top             =   3975
      Width           =   9780
      _cx             =   17251
      _cy             =   6879
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   4210752
      BackColorSel    =   16445674
      ForeColorSel    =   4210752
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   7
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdEarningsDeduct.frx":0420
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
      SubtotalPosition=   0
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      Begin TDBNumber6Ctl.TDBNumber txtAmount 
         Height          =   270
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   476
         Calculator      =   "frmAdEarningsDeduct.frx":0552
         Caption         =   "frmAdEarningsDeduct.frx":0572
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEarningsDeduct.frx":05DE
         Keys            =   "frmAdEarningsDeduct.frx":05FC
         Spin            =   "frmAdEarningsDeduct.frx":0646
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
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0.00"
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
   End
End
Attribute VB_Name = "frmAdEarningsDeduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim edKey               As Long
    Dim edRow               As Long
    Dim edCol               As Long

Private Sub cmdGenerate_Click()

    Dim rsOtherEarnings     As ADODB.Recordset
    Dim rsOtherDeductions   As ADODB.Recordset
    Dim rsTmp               As ADODB.Recordset
    
    Dim mCheck              As Boolean
    Dim mAddCol             As Boolean
    
    Dim I                   As Long
    
    Dim mQuery              As String
    
    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount <= 0 Then
        MsgBox "Please select a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If

    cmdGenerate.Enabled = False
    cmdSave.Enabled = False

    mCheck = False
    
    With fgEarningsDeduct
        .Cols = .ColIndex("fullname") + 1
        .Rows = 3
    End With

    NetOpen rsTmp, "select otherearningscode,otherearnings from otherearnings order by otherearnings"

    If rsTmp.RecordCount > 0 Then
        
        mAddCol = True
        
        With fgEarningsDeduct
            
            mCheck = True
            
            .Cols = .ColIndex("fullname") + 1
            
            rsTmp.MoveFirst
            
            Do While Not rsTmp.EOF
            
                .Cols = .Cols + 1
                .FixedAlignment(.Cols - 1) = flexAlignCenterCenter
                .ColAlignment(.Cols - 1) = flexAlignRightCenter
                .ColFormat(.Cols - 1) = "#,##0.00"
                .TextMatrix(0, .Cols - 1) = rsTmp!OtherEarningscode
                .TextMatrix(1, .Cols - 1) = "earnings"
                .TextMatrix(2, .Cols - 1) = UCase(rsTmp!OtherEarnings)
                .ColWidth(.Cols - 1) = 1700
            
                rsTmp.MoveNext
            
            Loop
            
        End With
        
    End If
    
    NetOpen rsTmp, "select otherdeductionscode,otherdeductions from otherdeductions order by otherdeductions"
    
    If rsTmp.RecordCount > 0 Then
        
        mAddCol = True
        
        With fgEarningsDeduct
            If mCheck = False Then
                .Cols = .ColIndex("fullname") + 1
            End If
            mCheck = True
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
            
                .Cols = .Cols + 1
                .FixedAlignment(.Cols - 1) = flexAlignCenterCenter
                .ColAlignment(.Cols - 1) = flexAlignRightCenter
                .ColFormat(.Cols - 1) = "#,##0.00"
                .TextMatrix(0, .Cols - 1) = rsTmp!OtherDeductionscode
                .TextMatrix(1, .Cols - 1) = "deductions"
                .TextMatrix(2, .Cols - 1) = UCase(rsTmp!OtherDeductions)
                .ColWidth(.Cols - 1) = 1700
            
                rsTmp.MoveNext
            
            Loop
        End With
    End If
    
    If mAddCol Then
        fgEarningsDeduct.Cols = fgEarningsDeduct.Cols + 1
    End If
    If mCheck = True Then
    
        If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
            If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
                If Trim(tdbCostCenter.Text) <> "" And Not IsNull(tdbCostCenter.SelectedItem) And tdbCostCenter.ApproxCount > 0 Then
                    If Trim(tdbSection.Text) <> "" And Not IsNull(tdbSection.SelectedItem) And tdbSection.ApproxCount > 0 Then
                        mQuery = "branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and " & _
                                 "costcentercode = '" & tdbCostCenter.BoundText & "' and sectioncode = '" & tdbSection.BoundText & "' and "
                    Else
                        mQuery = "branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and " & _
                                 "costcentercode = '" & tdbCostCenter.BoundText & "' and "
                    End If
                Else
                    mQuery = "branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and "
                End If
            Else
                mQuery = "branchcode = '" & tdbBranch.BoundText & "' and "
            End If
        End If
        
        NetOpen rsTmp, "select employeecode,concat(lastname,', ',firstname,' ',middlename) fullname, " & _
                                               "sectioncode,costcentercode,divisioncode,branchcode from employee " & _
                                               "where " & mQuery & " isactive = 'Y' and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' order by concat(lastname,', ',firstname,' ',middlename)"
        
        If rsTmp.RecordCount > 0 Then
            
            With fgEarningsDeduct
                .RowHidden(2) = False
                .Rows = 3
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("group")) = "A"
                    .TextMatrix(.Rows - 1, .ColIndex("branchcode")) = rsTmp!branchcode
                    .TextMatrix(.Rows - 1, .ColIndex("divisioncode")) = rsTmp!divisioncode
                    .TextMatrix(.Rows - 1, .ColIndex("costcentercode")) = rsTmp!costcentercode
                    .TextMatrix(.Rows - 1, .ColIndex("sectioncode")) = rsTmp!sectioncode
                    .TextMatrix(.Rows - 1, .ColIndex("employeecode")) = rsTmp!employeecode
                    .TextMatrix(.Rows - 1, .ColIndex("fullname")) = rsTmp!fullname
                    
                    NetOpen rsOtherEarnings, "select * from earnings where employeecode = '" & rsTmp!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "' order by otherearningscode"
                    
                    If rsOtherEarnings.RecordCount > 0 Then
                        rsOtherEarnings.MoveFirst
                        Do While Not rsOtherEarnings.EOF
                            I = 3
                            For I = 3 To .Cols - 1
                                If Trim(.TextMatrix(0, I)) = rsOtherEarnings!OtherEarningscode And Trim(.TextMatrix(1, I)) = "earnings" Then
                                    .TextMatrix(.Rows - 1, I) = rsOtherEarnings!amount
                                    Exit For
                                End If
                            Next
                            rsOtherEarnings.MoveNext
                        Loop
                    End If
                    
                    NetOpen rsOtherDeductions, "select * from deductions where employeecode = '" & rsTmp!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "' order by otherdeductionscode"
                    
                    If rsOtherDeductions.RecordCount > 0 Then
                        rsOtherDeductions.MoveFirst
                        Do While Not rsOtherDeductions.EOF
                            I = 3
                            For I = 3 To .Cols - 1
                                If Trim(.TextMatrix(0, I)) = rsOtherDeductions!OtherDeductionscode And Trim(.TextMatrix(1, I)) = "deductions" Then
                                    .TextMatrix(.Rows - 1, I) = rsOtherDeductions!amount
                                    Exit For
                                End If
                            Next
                            rsOtherDeductions.MoveNext
                        Loop
                    End If
                    
                    rsTmp.MoveNext
                Loop
                
                I = .ColIndex("fullname") + 1
                
                For I = I To .Cols - 1
                    .Subtotal flexSTSum, .ColIndex("group"), I, , &HE0E0E0, , True, "Total", , True
                Next
                .TextMatrix(.Rows - 1, .Cols - 1) = ""
            End With
            
        Else
            fgEarningsDeduct.RowHidden(2) = True
            MsgBox "No record found.", vbExclamation + vbOKOnly
        End If
        
        
        
    Else
        MsgBox "No record found.", vbExclamation + vbOKOnly
    End If
    
    cmdGenerate.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdSave_Click()
    
    Dim mRow            As Long
    Dim mCol            As Long
    
    Dim mWasSaved       As Boolean
    
    
    mWasSaved = False
    
    With fgEarningsDeduct
    
        If .Cols > 3 And .Rows > 3 Then
            
            cmdGenerate.Enabled = False
            cmdSave.Enabled = False
            Me.MousePointer = vbHourglass
            
            ConMain.Execute "set autocommit = 0"
            ConMain.BeginTrans
            
            mRow = 3
            
            For mRow = 3 To .Rows - 2
                
                mCol = .ColIndex("fullname") + 1
                
                ConMain.Execute "delete from earnings where employeecode = '" & .TextMatrix(mRow, .ColIndex("employeecode")) & "' and " & _
                                      "percode = '" & tdbPayrollPeriod.BoundText & "'"
                                      
                ConMain.Execute "delete from deductions where employeecode = '" & .TextMatrix(mRow, .ColIndex("employeecode")) & "' and " & _
                                      "percode = '" & tdbPayrollPeriod.BoundText & "'"
                
                For mCol = mCol To .Cols - 1
                    If IsNumeric(.TextMatrix(mRow, mCol)) Then
                    
                        mWasSaved = True
                        
                        If Trim(.TextMatrix(1, mCol)) = "earnings" Then
                        
                            ConMain.Execute "insert into earnings (otherearningscode,percode,employeecode,costcentercode, " & _
                                        "divisioncode,branchcode,sectioncode,payyear,paymonth,amount,remarks) values ( " & _
                                        "'" & .TextMatrix(0, mCol) & "','" & tdbPayrollPeriod.BoundText & "','" & .TextMatrix(mRow, .ColIndex("employeecode")) & "', " & _
                                        "'" & .TextMatrix(mRow, .ColIndex("costcentercode")) & "','" & .TextMatrix(mRow, .ColIndex("divisioncode")) & "', " & _
                                        "'" & .TextMatrix(mRow, .ColIndex("branchcode")) & "','" & .TextMatrix(mRow, .ColIndex("sectioncode")) & "', " & _
                                        "'" & tdbPayrollPeriod.Columns("payyear").Text & "','" & tdbPayrollPeriod.Columns("paymonth").Text & "','" & Format(.TextMatrix(mRow, mCol)) & "','')"
                        
                        Else
                            
                            ConMain.Execute "insert into deductions (otherdeductionscode,percode,employeecode,costcentercode, " & _
                                        "divisioncode,branchcode,sectioncode,payyear,paymonth,amount,remarks) values ( " & _
                                        "'" & .TextMatrix(0, mCol) & "','" & tdbPayrollPeriod.BoundText & "','" & .TextMatrix(mRow, .ColIndex("employeecode")) & "', " & _
                                        "'" & .TextMatrix(mRow, .ColIndex("costcentercode")) & "','" & .TextMatrix(mRow, .ColIndex("divisioncode")) & "', " & _
                                        "'" & .TextMatrix(mRow, .ColIndex("branchcode")) & "','" & .TextMatrix(mRow, .ColIndex("sectioncode")) & "', " & _
                                        "'" & tdbPayrollPeriod.Columns("payyear").Text & "','" & tdbPayrollPeriod.Columns("paymonth").Text & "','" & Format(.TextMatrix(mRow, mCol)) & "','')"
                        
                        End If
                    End If
                Next
            Next
            
            ConMain.CommitTrans
            
            cmdGenerate.Enabled = True
            cmdSave.Enabled = True
            Me.MousePointer = vbDefault
            
        End If
        
    End With
    
    If mWasSaved = True Then
        MsgBox "Process complete.", vbInformation + vbOKOnly
    End If
    
End Sub

Private Sub fgEarningsDeduct_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With fgEarningsDeduct
        If Col = .Cols - 1 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub fgEarningsDeduct_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Cancel = True
    
    With fgEarningsDeduct
        
        If Trim(.TextMatrix(.Row, .ColIndex("employeecode"))) <> "" And Col > .ColIndex("fullname") Then
            
            ' save row/col being edited
            edRow = fgEarningsDeduct.Row
            edCol = fgEarningsDeduct.Col
            
            ' move editor over cell
            txtAmount.Move fgEarningsDeduct.CellLeft, fgEarningsDeduct.CellTop, fgEarningsDeduct.CellWidth - Screen.TwipsPerPixelX, fgEarningsDeduct.CellHeight - Screen.TwipsPerPixelY
            
            ' initialize editor
            txtAmount.Text = IIf(IsNumeric(fgEarningsDeduct.Text), Format(fgEarningsDeduct.Text, "#,##0.00"), "0.00")
            
            If edKey > 32 Then
                If IsNumeric(Chr(edKey)) Then
                    txtAmount.Text = Format(Chr(edKey), "#,##0.00")
                Else
                    txtAmount.Text = "0.00"
                End If
                txtAmount.SelStart = 1
            ElseIf edKey = 32 Then
                txtAmount.SelStart = 32000
            Else
                txtAmount.SelStart = 0
                txtAmount.SelLength = 32000
            End If
            
            ' activate editor
            txtAmount.Visible = True
            txtAmount.SetFocus
        
        End If
    
    End With
    
End Sub

Private Sub fgEarningsDeduct_KeyPress(KeyAscii As Integer)

    edKey = KeyAscii
    
End Sub

Private Sub fgEarningsDeduct_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    edKey = 0

End Sub

Private Sub fgEarningsDeduct_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    
    If txtAmount.Visible Then
        txtAmount.Move fgEarningsDeduct.CellLeft, fgEarningsDeduct.CellTop, fgEarningsDeduct.CellWidth - Screen.TwipsPerPixelX, fgEarningsDeduct.CellHeight - Screen.TwipsPerPixelY
    End If
    
End Sub

Private Sub fgEarningsDeduct_AfterUserResize(ByVal Row As Long, ByVal Col As Long)

    If txtAmount.Visible Then
        txtAmount.Move fgEarningsDeduct.CellLeft, fgEarningsDeduct.CellTop, fgEarningsDeduct.CellWidth - Screen.TwipsPerPixelX, fgEarningsDeduct.CellHeight - Screen.TwipsPerPixelY
    End If

End Sub


' -- handle key press on the editor
Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   
    
    With fgEarningsDeduct
    
        Select Case KeyCode
        
            Case vbKeyRight
                If txtAmount.SelStart = Len(txtAmount.Text) Then
                    .SetFocus
                    .Col = .Col + 1
                End If
            Case vbKeyLeft
                If txtAmount.SelStart = Len(txtAmount.Text) Then
                    .SetFocus
                    .Col = .Col - 1
                End If
            Case vbKeyReturn
                
                .SetFocus
                .Row = .Row + 1
            Case vbKeyEscape
                txtAmount.Text = Format(.TextMatrix(edRow, edCol), "#,##0.00")
                .SetFocus
        End Select
        
    End With
    
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13, 27:
            KeyAscii = 0
    End Select
End Sub

Private Sub txtAmount_LostFocus()

    Dim I As Long

    With fgEarningsDeduct
        
        If IsNumeric(txtAmount.Text) Then
            .TextMatrix(edRow, edCol) = IIf(CDbl(txtAmount.Text) = 0, "", txtAmount.Text)
        Else
            .TextMatrix(edRow, edCol) = ""
        End If
        txtAmount.Visible = False
        edKey = 0
        
        I = .ColIndex("fullname") + 1
                    
        For I = I To .Cols - 1
            .Subtotal flexSTSum, .ColIndex("group"), I, , &HE0E0E0, , True, "Total", , True
        Next
        .TextMatrix(.Rows - 1, .Cols - 1) = ""
    End With
    
End Sub

Private Sub Form_Load()

    Add_MDIButton Me.Name, titlebar.Caption
    
    bind_tdb ConMain, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto,payfreqcode,payyear,paymonth from payrollperiod", "description", "percode"
    
    bind_tdb ConMain, tdbBranch, "select branchcode, branch from branch order by branch", "branch", "branchcode"
    
    With fgEarningsDeduct
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColHidden(.ColIndex("group")) = True
        .ColHidden(.ColIndex("branchcode")) = True
        .ColHidden(.ColIndex("divisioncode")) = True
        .ColHidden(.ColIndex("costcentercode")) = True
        .ColHidden(.ColIndex("sectioncode")) = True
        .ColHidden(.ColIndex("employeecode")) = True
    End With
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Remove_MDIButton Me.Name
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    titlebar.Move 0, 0, Me.ScaleWidth
    
    With fra1
      .Top = titlebar.Height + titlebar.Top
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With fgEarningsDeduct
        .Top = fra1.Top + fra1.Height
        .Left = 0
        .Height = Me.ScaleHeight - .Top
        .Width = Me.ScaleWidth
    End With

End Sub

Private Sub tdbBranch_GotFocus()
    
    tdbBranch.Tag = tdbBranch.BoundText
    
End Sub

Private Sub tdbBranch_LostFocus()
    
    If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
        If (tdbBranch.Tag <> tdbBranch.BoundText) Or (Trim(tdbDivision.Text) = "" Or IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount <= 0) Then
        
            Set tdbDivision.RowSource = Nothing
            Set tdbCostCenter.RowSource = Nothing
            Set tdbSection.RowSource = Nothing
        
            bind_tdb ConMain, tdbDivision, "select divisioncode,division from division where branchcode = '" & tdbBranch.BoundText & "' order by division", "division", "divisioncode"
            tdbDivision.BoundText = ""
            tdbCostCenter.BoundText = ""
            tdbSection.BoundText = ""
                
        End If
    Else
        
        tdbDivision.BoundText = ""
        tdbCostCenter.BoundText = ""
        tdbSection.BoundText = ""
        
        Set tdbDivision.RowSource = Nothing
        Set tdbCostCenter.RowSource = Nothing
        Set tdbSection.RowSource = Nothing
        
    End If
    
End Sub

Private Sub tdbCostCenter_GotFocus()
    tdbCostCenter.Tag = tdbCostCenter.BoundText
End Sub

Private Sub tdbCostCenter_LostFocus()

    If Trim(tdbCostCenter.Text) <> "" And Not IsNull(tdbCostCenter.SelectedItem) And tdbCostCenter.ApproxCount > 0 Then
        If tdbCostCenter.Tag <> tdbCostCenter.BoundText Then
            
            Set tdbSection.RowSource = Nothing
            
            bind_tdb ConMain, tdbSection, "select sectioncode, sectionname from section " & _
                    "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and costcentercode = '" & tdbCostCenter.BoundText & "' order by sectionname", "sectionname", "sectioncode"
                    
            tdbSection.BoundText = ""
            
        End If
    Else
  
        tdbSection.BoundText = ""
        
        Set tdbSection.RowSource = Nothing
        
    End If

End Sub

Private Sub tdbDivision_GotFocus()
    
    tdbDivision.Tag = tdbDivision.BoundText
    
End Sub

Private Sub tdbDivision_LostFocus()
    
    If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
        If tdbDivision.Tag <> tdbDivision.BoundText Then
            
            Set tdbCostCenter.RowSource = Nothing
            Set tdbSection.RowSource = Nothing
            
            bind_tdb ConMain, tdbCostCenter, "select costcentercode,costcenter from costcenter where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' order by costcenter", "costcenter", "costcentercode"
            
            tdbCostCenter.BoundText = ""
            tdbSection.BoundText = ""
                        
        End If
    Else
    
        tdbCostCenter.BoundText = ""
        tdbSection.BoundText = ""
        
        Set tdbCostCenter.RowSource = Nothing
        Set tdbSection.RowSource = Nothing
        
    End If
    
End Sub

Private Sub txtSearch_Change()
    On Error Resume Next
    Static cur_loc As Integer
    Static lapos As Boolean
    Static Change_Row As Long
    fgEarningsDeduct.Select fgEarningsDeduct.FindRow(txtSearch.Text, 1, fgEarningsDeduct.ColIndex("fullname"), False, False), fgEarningsDeduct.ColIndex("fullname")
    fgEarningsDeduct.ShowCell fgEarningsDeduct.FindRow(txtSearch.Text, 1, fgEarningsDeduct.ColIndex("fullname"), False, False), fgEarningsDeduct.ColIndex("fullname")
    If Change_Row = CLng(fgEarningsDeduct.Row) Then
      If UCase(Mid(txtSearch.Text, 1, txtSearch.SelStart)) <> UCase(Mid(fgEarningsDeduct.TextMatrix(fgEarningsDeduct.Row, fgEarningsDeduct.ColIndex("fullname")), 1, txtSearch.SelStart)) Then
        If txtSearch.SelStart > 0 Then
        cur_loc = txtSearch.SelStart
        txtSearch.SelStart = cur_loc - 1
        End If
      End If
    End If
    If lapos = True Then
      cur_loc = txtSearch.SelStart
    End If
    lapos = False
    If Len(txtSearch.Text) = 1 Then cur_loc = 1
    txtSearch.Text = fgEarningsDeduct.TextMatrix(fgEarningsDeduct.Row, fgEarningsDeduct.ColIndex("fullname"))
    Change_Row = CLng(fgEarningsDeduct.Row)
    lapos = True
    txtSearch.SelStart = cur_loc
    txtSearch.SelLength = Len(txtSearch.Text)

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 8 Then
    If txtSearch.SelStart > 0 Then
      txtSearch.SelStart = txtSearch.SelStart - 1
      txtSearch.SelLength = Len(txtSearch.Text)
    End If
  End If
End Sub

