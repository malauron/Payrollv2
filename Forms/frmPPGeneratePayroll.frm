VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPGeneratePayroll 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Payroll"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   15
      TabIndex        =   5
      Top             =   -75
      Width           =   6105
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   105
         Left            =   1980
         TabIndex        =   6
         Top             =   2295
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   345
         Left            =   1980
         TabIndex        =   3
         Tag             =   "Municipal"
         Top             =   1680
         Width           =   4005
         _ExtentX        =   7064
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
         _PropDict       =   $"frmPPGeneratePayroll.frx":0000
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF8FAFA&,.bold=0,.fontsize=825"
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
         Left            =   1980
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   720
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   609
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   609
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
         _PropDict       =   $"frmPPGeneratePayroll.frx":00AA
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF8FAFA&,.bold=0,.fontsize=825"
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
         Left            =   1980
         TabIndex        =   2
         Tag             =   "Municipal"
         Top             =   1140
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   609
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   609
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
         _PropDict       =   $"frmPPGeneratePayroll.frx":0154
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF8FAFA&,.bold=0,.fontsize=825"
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
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   105
         Left            =   1980
         TabIndex        =   7
         Top             =   2430
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1980
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   180
         Width           =   4005
         _ExtentX        =   7064
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
         _PropDict       =   $"frmPPGeneratePayroll.frx":01FE
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
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   390
         Left            =   60
         TabIndex        =   4
         Top             =   2220
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   688
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
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPPGeneratePayroll.frx":02A8
         cBack           =   14737632
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   120
         X2              =   6000
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   120
         X2              =   6000
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         Height          =   195
         Index           =   2
         Left            =   465
         TabIndex        =   11
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
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
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee(s)"
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
         Left            =   210
         TabIndex        =   8
         Top             =   1740
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frmPPGeneratePayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NTA_Data
  mNTA_Desc As String
  mNTA_Amt  As Double
End Type
  
  
Dim rsEmployee            As ADODB.Recordset
Dim rsParmtr              As ADODB.Recordset

Private Sub Form_Load()
    
'    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
'    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
'    SendMessage pb2.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
'    SendMessage pb2.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    bind_tdb ConMain, tdbPayrollPeriod, "select percode,description,wrkdatefrom,wrkdateto,payfreqcode,payyear,paymonth from payrollperiod where fnlz <> 'Y' order by percode desc", "description", "percode"
    
    bind_tdb ConMain, tdbBranch, "select branchcode, branch from branch order by branch", "branch", "branchcode"
  
End Sub

Private Sub tdbBranch_ItemChange()

  bind_tdb ConMain, tdbDivision, "select divisioncode,division from division " & _
            "where branchcode = '" & tdbBranch.BoundText & "' order by division", "division", "divisioncode"
  
  If tdbDivision.ApproxCount > 0 Then
    tdbDivision.BoundText = tdbDivision.Columns(0).Text
  End If
  
End Sub

Private Sub tdbBranch_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbBranch, tdbBranch.RowSource, tdbBranch.Text
    tdbBranch_ItemChange
  End If
  
End Sub

Private Sub tdbDivision_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchList KeyAscii, tdbDivision, tdbDivision.RowSource, tdbDivision.Text
    End If
    
End Sub

Private Sub tdbDivision_Itemchange()

    tdbEmployee.BoundText = ""
    
    bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname, ', ', firstname,' ',middlename) fullname from employee " & _
                "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and isactive <> 'N' " & _
                "order by concat(lastname, ', ', firstname,' ',middlename) ", "fullname", "employeecode"
                
    
End Sub

Private Sub cmdGenerate_Click()

    Dim rsPayPeriod           As ADODB.Recordset
    Dim rsDTR                 As ADODB.Recordset
    Dim rsOT                  As ADODB.Recordset
    Dim rsLeave               As ADODB.Recordset
    Dim rsEarnings            As ADODB.Recordset
    Dim rsDeductions          As ADODB.Recordset
    Dim rsLoanDed             As ADODB.Recordset
    Dim rsLVLimit             As ADODB.Recordset
    Dim rsChk                 As ADODB.Recordset
    Dim rsTmp                 As ADODB.Recordset
    Dim rsPrevPeriod          As ADODB.Recordset
    
    Dim mFEdesc1              As String
    Dim mFEdesc2              As String
    Dim mFEdesc3              As String
    Dim mFEdesc4              As String
    Dim mFEdesc5              As String
    Dim mFEamnt1              As Double
    Dim mFEamnt2              As Double
    Dim mFEamnt3              As Double
    Dim mFEamnt4              As Double
    Dim mFEamnt5              As Double
    
    Dim mODdesc1              As String
    Dim mODdesc2              As String
    Dim mODdesc3              As String
    Dim mODdesc4              As String
    Dim mODdesc5              As String
    Dim mODdesc6              As String
    Dim mODdesc7              As String
    Dim mODamnt1              As Double
    Dim mODamnt2              As Double
    Dim mODamnt3              As Double
    Dim mODamnt4              As Double
    Dim mODamnt5              As Double
    Dim mODamnt6              As Double
    Dim mODamnt7              As Double
    
    Dim mLDdesc1              As String
    Dim mLDdesc2              As String
    Dim mLDdesc3              As String
    Dim mLDdesc4              As String
    Dim mLDdesc5              As String
    Dim mLDdesc6              As String
    Dim mLDdesc7              As String
    Dim mLDdesc8              As String
    Dim mLDamnt1              As Double
    Dim mLDamnt2              As Double
    Dim mLDamnt3              As Double
    Dim mLDamnt4              As Double
    Dim mLDamnt5              As Double
    Dim mLDamnt6              As Double
    Dim mLDamnt7              As Double
    Dim mLDamnt8              As Double
    Dim mLBal1                As Double
    Dim mLBal2                As Double
    Dim mLBal3                As Double
    Dim mLBal4                As Double
    Dim mLBal5                As Double
    Dim mLBal6                As Double
    Dim mLBal7                As Double
    
    Dim mLNoOfPay1            As Double
    Dim mLNoOfPay2            As Double
    Dim mLNoOfPay3            As Double
    Dim mLNoOfPay4            As Double
    Dim mLNoOfPay5            As Double
    Dim mLNoOfPay6            As Double
    Dim mLNoOfPay7            As Double
    
    Dim CTR                   As Integer
    Dim mPRCtr                As Integer
    Dim mDTRCtr               As Integer
    
    Dim mBasic                As Double
    Dim mGross                As Double
    Dim mNet                  As Double
    
    Dim mRstDays              As Double
    Dim mRstHrs               As Double
    Dim mRstAmnt              As Double
    
    Dim mLegDays              As Double
    Dim mLegHrs               As Double
    Dim mLegAmnt              As Double
    
    Dim mSpcDays              As Double
    Dim mSpcHrs               As Double
    Dim mSpcAmnt              As Double
    
    Dim mLvLimit              As Double
    Dim mLvWPDay              As Double
    Dim mLvWPHrs              As Double
    Dim mLvWPDays             As Double
    Dim mLvWPAmnt             As Double
    Dim mLvWoPDays            As Double
    Dim mLvWoPAmnt            As Double
    
    Dim mTtlWrkDays           As Integer
    Dim mTtlWrkHrs            As Double
    
    
    Dim mRegWrkHrs            As Double
    Dim mDaysWrk              As Double
    Dim mAbsDays              As Double
    Dim mLatehrs              As Double
    Dim mUtHrs                As Double
    
    Dim mAbsAmnt              As Double
    Dim mLateAmnt             As Double
    Dim mUtAmnt               As Double
    
    Dim mRegAmnt              As Double
    
    Dim mOtRegHrs             As Double
    Dim mOtRegAmnt            As Double
    
    Dim mNtDfHrs              As Double
    Dim mNtDfAmnt             As Double
    Dim mNtDfHrsReg           As Double
    Dim mNtDfAmntReg          As Double
    Dim mNtDfHrsLeg           As Double
    Dim mNtDfAmntLeg          As Double
    Dim mNtDfHrsSpc           As Double
    Dim mNtDfAmntSpc          As Double
        
    Dim mEarnings             As Double
    Dim mDeductions           As Double
    Dim mLoanded              As Double
    Dim mNonTaxAllowBasic     As Double
    Dim mNonTaxAllowNet       As Double
    
    Dim mTtlMealAllow         As Double
    Dim mFixedDed             As Double
    
    Dim mFixedEarnings        As Double
    Dim mCola                 As Double
    
    Dim mSSSBasic             As Double
    Dim regER                 As Double
    Dim regEE                 As Double
    Dim ecER                  As Double
    Dim ecEE                  As Double
    Dim mpfER                 As Double
    Dim mpfEE                 As Double
    Dim ttlER                 As Double
    Dim ttlEE                 As Double
    
    Dim philBasic             As Double
    Dim mPhilAmnt             As Double
    Dim mPhilEr               As Double
    
    Dim mHdmfAmnt             As Double
    Dim mHdmfEr               As Double
    Dim mTaxableInc           As Double
    Dim mTaxAmnt              As Double
    
    Dim mCompHol        As Boolean

     If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        MsgBox "Please select a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If
    
    NetOpen rsParmtr, "select * from parmtr"
    If rsParmtr.RecordCount <= 0 Then
        MsgBox "Parameters are not set.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    NetOpen rsPrevPeriod, "SELECT COUNT(percode) ctr FROM payrollperiod WHERE percode < " & tdbPayrollPeriod.BoundText & " AND fnlz = 'N'"
    
    If rsPrevPeriod.RecordCount > 0 Then
      If rsPrevPeriod.Fields("ctr") > 0 Then
        MsgBox "Please finalize the previous payroll before continuing to generate the current payroll period.", vbExclamation + vbOKOnly
        Exit Sub
      End If
    End If
    
    If MsgBox("Confirm generate payroll.", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
        If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
                If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
                    NetOpen rsEmployee, "select x1.*,x2.ttl_days from employee x1 " & _
                                        "left outer join wrkdays x2 on x1.wrkdays_id=x2.wrkdays_id " & _
                                        "where x1.employeecode  = '" & tdbEmployee.BoundText & "' and x1.employeecode in (" & _
                                        "select employeecode from dtr where employeecode = " & tdbEmployee.BoundText & ") and " & _
                                        "x1.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                Else
                    NetOpen rsEmployee, "select x2.*,x3.ttl_days from (select employeecode from dtr where percode = " & tdbPayrollPeriod.BoundText & ") x1 " & _
                                        "left outer join employee x2 on x1.employeecode = x2.employeecode  " & _
                                        "left outer join wrkdays x3 on x2.wrkdays_id=x3.wrkdays_id " & _
                                        "where x2.branchcode = '" & tdbBranch.BoundText & "' and " & _
                                        "x2.divisioncode = '" & tdbDivision.BoundText & "' and " & _
                                        "x2.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                End If
        Else
            NetOpen rsEmployee, "select x2.*,x3.ttl_days from (select employeecode from dtr where percode = " & tdbPayrollPeriod.BoundText & ") x1 " & _
                                        "left outer join employee x2 on x1.employeecode = x2.employeecode  " & _
                                        "left outer join wrkdays x3 on x2.wrkdays_id=x3.wrkdays_id " & _
                                        "where x2.branchcode = '" & tdbBranch.BoundText & "' and " & _
                                        "x2.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
        End If
    Else
        NetOpen rsEmployee, "select x2.*,x3.ttl_days from (select employeecode from dtr where percode = " & tdbPayrollPeriod.BoundText & ") x1 " & _
                                        "left outer join employee x2 on x1.employeecode = x2.employeecode  " & _
                                        "left outer join wrkdays x3 on x2.wrkdays_id=x3.wrkdays_id " & _
                                        "where x2.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
    End If
   
    If rsEmployee.RecordCount > 0 Then
        
        NetOpen rsPayPeriod, "select * from payrollperiod where percode = " & tdbPayrollPeriod.BoundText & ""
    
        fra1.Enabled = False
        Me.MousePointer = vbHourglass
        cmdGenerate.Enabled = False
        
        pb1.Max = rsEmployee.RecordCount
        pb1.Value = 0
        
        rsEmployee.MoveFirst
        
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        
        
        If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
            If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
                    If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
                        ConMain.Execute "Delete from payroll where percode = " & tdbPayrollPeriod.BoundText & " and " & _
                                        "employeecode = " & tdbEmployee.BoundText & ""
                    Else
                        ConMain.Execute "Delete from payroll where (percode = " & tdbPayrollPeriod.BoundText & " and " & _
                                        "divisioncode = " & tdbDivision.BoundText & " and branchcode = " & tdbBranch.BoundText & ") or " & _
                                        "(percode = " & tdbPayrollPeriod.BoundText & " and " & _
                                        "employeecode in  (select employeecode from employee where divisioncode = " & tdbDivision.BoundText & " and branchcode = " & tdbBranch.BoundText & "))"
                    End If
            Else
                ConMain.Execute "Delete from payroll where (percode = " & tdbPayrollPeriod.BoundText & " and " & _
                                        "branchcode = " & tdbBranch.BoundText & ") or " & _
                                        "(percode = " & tdbPayrollPeriod.BoundText & " and " & _
                                        "employeecode in  (select employeecode from employee where branchcode = " & tdbBranch.BoundText & "))"
            End If
        Else
            ConMain.Execute "delete from payroll where percode = " & tdbPayrollPeriod.BoundText & ""
        End If
        
        
        Do While Not rsEmployee.EOF
        
            pb1.Value = pb1.Value + 1
            
            mFEdesc1 = ""
            mFEdesc2 = ""
            mFEdesc3 = ""
            mFEdesc4 = ""
            mFEdesc5 = ""
            mFEamnt1 = 0
            mFEamnt2 = 0
            mFEamnt3 = 0
            mFEamnt4 = 0
            mFEamnt5 = 0
            
            mODdesc1 = ""
            mODdesc2 = ""
            mODdesc3 = ""
            mODdesc4 = ""
            mODdesc5 = ""
            mODdesc6 = ""
            mODdesc7 = ""
            mODamnt1 = 0
            mODamnt2 = 0
            mODamnt3 = 0
            mODamnt4 = 0
            mODamnt5 = 0
            mODamnt6 = 0
            mODamnt7 = 0
            
            mLDdesc1 = ""
            mLDdesc2 = ""
            mLDdesc3 = ""
            mLDdesc4 = ""
            mLDdesc5 = ""
            mLDdesc6 = ""
            mLDdesc7 = ""
            mLDdesc8 = ""
            mLDamnt1 = 0
            mLDamnt2 = 0
            mLDamnt3 = 0
            mLDamnt4 = 0
            mLDamnt5 = 0
            mLDamnt6 = 0
            mLDamnt7 = 0
            mLDamnt8 = 0
            mLBal1 = 0
            mLBal2 = 0
            mLBal3 = 0
            mLBal4 = 0
            mLBal5 = 0
            mLBal6 = 0
            mLBal7 = 0
            mLNoOfPay1 = 0
            mLNoOfPay2 = 0
            mLNoOfPay3 = 0
            mLNoOfPay4 = 0
            mLNoOfPay5 = 0
            mLNoOfPay6 = 0
            mLNoOfPay7 = 0
            
            
            mBasic = 0
            mGross = 0
            mNet = 0
            
            mLvLimit = 0
            mLvWPDay = 0
            mLvWPDays = 0
            mLvWPAmnt = 0
            
            mRstDays = 0
            mRstHrs = 0
            mRstAmnt = 0
            mLegDays = 0
            mLegHrs = 0
            mLegAmnt = 0
            mSpcDays = 0
            mSpcHrs = 0
            mSpcAmnt = 0
            
            mRegWrkHrs = 0
            mDaysWrk = 0
            mAbsDays = 0
            mLatehrs = 0
            mUtHrs = 0
            
            mAbsAmnt = 0
            mLateAmnt = 0
            mUtAmnt = 0
            
            mRegAmnt = 0
            
            mOtRegHrs = 0
            mOtRegAmnt = 0
            
            mNtDfHrs = 0
            mNtDfAmnt = 0
            mNtDfHrsReg = 0
            mNtDfAmntReg = 0
            mNtDfHrsLeg = 0
            mNtDfAmntLeg = 0
            mNtDfHrsSpc = 0
            mNtDfAmntSpc = 0
            
            mEarnings = 0
            mDeductions = 0
            mLoanded = 0
            mNonTaxAllowBasic = 0
            mNonTaxAllowNet = 0
            
            mTtlMealAllow = 0
            mFixedDed = 0
            
            mFixedEarnings = 0
            mCola = 0
            
            mSSSBasic = 0
            regER = 0
            regEE = 0
            ecER = 0
            ecEE = 0
            mpfER = 0
            mpfEE = 0
            ttlEE = 0
            ttlER = 0
            
            philBasic = 0
            mPhilAmnt = 0
            mPhilEr = 0
            
            mHdmfAmnt = 0
            mHdmfEr = 0
            mTaxableInc = 0
            mTaxAmnt = 0
            
            mTtlWrkDays = rsEmployee!ttl_days
            mTtlWrkHrs = mTtlWrkDays * 8
            
            If rsEmployee!ratetypecode = 4 Then  'monthly rate employees
            
                'mRegAmnt = rsEmployee!Hourly_Rate * 104
                mRegAmnt = rsEmployee!monthly_rate / 2
                
                'mRegWrkHrs = 104
                mRegWrkHrs = mTtlWrkHrs / 2
                
                'mTtlMealAllow = rsEmployee!mealallow2 * 13
                mTtlMealAllow = rsEmployee!mealallow2 * mTtlWrkDays / 2
                
                'mFixedDed = rsEmployee!MealAllow * 13
                mFixedDed = rsEmployee!MealAllow * mTtlWrkDays / 2
                
                NetOpen rsDTR, "select * from dtr where employeecode = " & rsEmployee!employeecode & " and percode = " & tdbPayrollPeriod.BoundText & ""
                
                If rsDTR.RecordCount > 0 Then
                    mAbsDays = rsDTR!absdays
                    mAbsAmnt = (rsEmployee!daily_rate * rsDTR!absdays)
                    mLateAmnt = (rsEmployee!Hourly_Rate * rsDTR!late)
                    mUtAmnt = (rsEmployee!Hourly_Rate * rsDTR!undertime)
                    mLegAmnt = (rsEmployee!Hourly_Rate * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDTR!legdays)
                    mSpcAmnt = (rsEmployee!Hourly_Rate * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDTR!spcdays)
                    mRstAmnt = (rsEmployee!Hourly_Rate * rsDTR!restdays)
                    mNtDfAmntReg = (rsEmployee!Hourly_Rate * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffReg)
                    mNtDfAmntLeg = (rsEmployee!Hourly_Rate * Format(rsParmtr!nitelegprct / 100, "#,##0.00") * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffLeg)
                    mNtDfAmntSpc = (rsEmployee!Hourly_Rate * Format(rsParmtr!nitespcprct / 100, "#,##0.00") * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffSpc)
                    mNtDfAmnt = mNtDfAmntReg + mNtDfAmntLeg + mNtDfAmntSpc
                    
                    mLatehrs = rsDTR!late
                    mUtHrs = rsDTR!undertime
                    mLegHrs = rsDTR!legdays
                    mSpcHrs = rsDTR!spcdays
                    mRstHrs = rsDTR!restdays
                    mNtDfHrsReg = rsDTR!nightdiffReg
                    mNtDfHrsLeg = rsDTR!nightdiffLeg
                    mNtDfHrsSpc = rsDTR!nightdiffSpc
                    mNtDfHrs = mNtDfHrsReg + mNtDfHrsLeg + mNtDfHrsSpc
                End If
                
                If rsEmployee!FixedEarnings > 0 Then
                                      '********Hours Worked*********     *****absent hrs + late + undertime ******
                    mFixedEarnings = (rsEmployee!FixedEarnings / mTtlWrkHrs) * ((mTtlWrkHrs / 2) - ((mAbsDays * 8) + mLatehrs + mUtHrs))
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDTR!legdays)
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDTR!spcdays)
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * rsDTR!restdays)
                End If
                
                If rsEmployee!cola > 0 Then
                    mCola = (rsEmployee!cola / 2) - (rsEmployee!cola / mTtlWrkDays * mAbsDays)
                    If mCola < 0 Then mCola = 0
                End If
                
            ElseIf rsEmployee!ratetypecode = 1 Then   'daily rate employee
                
                NetOpen rsDTR, "select * from dtr where employeecode = " & rsEmployee!employeecode & " and percode = " & tdbPayrollPeriod.BoundText & ""
                
                If rsDTR.RecordCount > 0 Then
                
                    mDaysWrk = rsDTR!dayswork
                    mTtlMealAllow = rsDTR!dayswork * rsEmployee!mealallow2
                    mFixedDed = rsDTR!dayswork * rsEmployee!MealAllow
                    mRegAmnt = rsEmployee!daily_rate * rsDTR!dayswork
                    mLateAmnt = (rsEmployee!Hourly_Rate * rsDTR!late)
                    mUtAmnt = (rsEmployee!Hourly_Rate * rsDTR!undertime)
                    mLegAmnt = (rsEmployee!Hourly_Rate * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDTR!legdays)
                    mSpcAmnt = (rsEmployee!Hourly_Rate * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDTR!spcdays)
                    mRstAmnt = (rsEmployee!Hourly_Rate * rsDTR!restdays)
                    'mNtDfAmnt = (rsEmployee!Hourly_Rate * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiff)
                    mNtDfAmntReg = (rsEmployee!Hourly_Rate * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffReg)
                    mNtDfAmntLeg = (rsEmployee!Hourly_Rate * Format(rsParmtr!nitelegprct / 100, "#,##0.00") * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffLeg)
                    mNtDfAmntSpc = (rsEmployee!Hourly_Rate * Format(rsParmtr!nitespcprct / 100, "#,##0.00") * Format(rsParmtr!niteregprct / 100, "#,##0.00") * rsDTR!nightdiffSpc)
                    mNtDfAmnt = mNtDfAmntReg + mNtDfAmntLeg + mNtDfAmntSpc
                    
                    mLatehrs = rsDTR!late
                    mUtHrs = rsDTR!undertime
                    mLegHrs = rsDTR!legdays
                    mSpcHrs = rsDTR!spcdays
                    mRstHrs = rsDTR!restdays
                    'mNtDfHrs = rsDTR!nightdiff
                    mNtDfHrsReg = rsDTR!nightdiffReg
                    mNtDfHrsLeg = rsDTR!nightdiffLeg
                    mNtDfHrsSpc = rsDTR!nightdiffSpc
                    mNtDfHrs = mNtDfHrsReg + mNtDfHrsLeg + mNtDfHrsSpc
                    
                End If
                
                If rsEmployee!FixedEarnings > 0 Then
                                      '*********hourly rate**********    **hrs worked**      *********hourly rate**********     *no. of hrs tardy*
                    mFixedEarnings = ((rsEmployee!FixedEarnings / mTtlWrkHrs) * (mDaysWrk * 8)) - ((rsEmployee!FixedEarnings / mTtlWrkHrs) * (mLatehrs + mUtHrs))
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDTR!legdays)
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDTR!spcdays)
                    mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkHrs) * rsDTR!restdays)
                End If
                
                If rsEmployee!cola > 0 Then
                            '*********hourly rate**********    **hrs worked**      *********hourly rate**********     *no. of hrs tardy*
                    mCola = rsEmployee!cola / mTtlWrkDays * mDaysWrk
                End If
                
            End If
            
            CompNonTaxAllow rsDTR, mTtlWrkHrs, mDaysWrk, mAbsDays, mLatehrs, mUtHrs, mNonTaxAllowBasic, mNonTaxAllowNet
            
            NetOpen rsOT, "select sum(othrs) ttlothrs from overtimelne where fnlz <> 'Y' and employeecode = '" & rsEmployee!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "' and status = 'Approved'"
            
            With rsOT
                If .RecordCount > 0 Then
                    If Not IsNull(!ttlothrs) Then
                        mOtRegHrs = !ttlothrs
                        mOtRegAmnt = mOtRegHrs * rsEmployee!Hourly_Rate * 1.25
                    End If
                End If
            End With
                       
            CTR = 1
            ConMain.Execute "delete from lvavailed where percode=" & tdbPayrollPeriod.BoundText & " And employeecode = " & rsEmployee!employeecode
            ConMain.Execute "delete from lvavailed_payslips where percode=" & tdbPayrollPeriod.BoundText & " And employeecode = " & rsEmployee!employeecode
            
            NetOpen rsLeave, "SELECT x1.leavetypescode,x1.lvhrs,IFNULL(x2.lvlimithrs,0) lvlimithrs,X3.DESCRIPTION FROM " & _
                             "(SELECT leavetypescode,SUM(leaveapp_hours*((firstshift*0.5)+(secondshift*0.5))) lvhrs FROM leaveapp_lines " & _
                             "WHERE employeecode=" & rsEmployee!employeecode & " AND withpay=1 AND (leaveapp_date BETWEEN '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' AND  '" & Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "YYYY-MM-DD") & "')  " & _
                             "GROUP BY leavetypescode " & _
                             "ORDER BY leavetypescode) x1 " & _
                             "LEFT OUTER JOIN (SELECT *,(lvlimit*" & rsParmtr!leavehrsperday & ") lvlimithrs FROM lvlimit WHERE employeecode = " & rsEmployee!employeecode & " AND payyear = '" & tdbPayrollPeriod.Columns("payyear").Text & "' ) x2 ON x1.leavetypescode=x2.leavetypescode " & _
                             "LEFT OUTER JOIN LEAVETYPES X3 ON X1.LEAVETYPESCODE=X3.LEAVETYPESCODE"
                             
            With rsLeave
              If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                
                  mLvWPHrs = 0
                  mLvWPDay = 0
                  mLvLimit = 0
                  
                  If CDbl(!lvlimithrs) > 0 Then
                    If CDbl(!lvlimithrs) > CDbl(!lvhrs) Then
                      mLvWPHrs = CDbl(!lvhrs)
                      mLvWPDay = CDbl(!lvhrs) / rsParmtr!leavehrsperday
                      mLvLimit = (CDbl(!lvlimithrs) - CDbl(!lvhrs)) / rsParmtr!leavehrsperday
                    Else
                      mLvWPHrs = CDbl(!lvlimithrs)
                      mLvWPDay = CDbl(!lvlimithrs) / rsParmtr!leavehrsperday
                      mLvLimit = 0
                      mLvWoPDays = mLvWoPDays + ((CDbl(!lvhrs) - CDbl(!lvlimithrs)) / rsParmtr!leavehrsperday)
                    End If
                    
                    If CTR = 1 Then
                      ConMain.Execute "insert into lvavailed_payslips (percode,employeecode,desc1,limit1,days1,amnt1) values (" & tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & ",'" & CStr(!Description) & "'," & mLvLimit & "," & mLvWPDay & "," & mLvWPDay * rsEmployee!daily_rate & ")"
                    ElseIf CTR > 1 And CTR < 4 Then
                      ConMain.Execute "update lvavailed_payslips set desc" & CTR & "='" & CStr(!Description) & "',limit" & CTR & "=" & mLvLimit & ",days" & CTR & "=" & mLvWPDay & ",amnt" & CTR & "=" & mLvWPDay * rsEmployee!daily_rate & " where percode=" & tdbPayrollPeriod.BoundText & " and employeecode=" & rsEmployee!employeecode & ""
                    Else
                      ConMain.Execute "update lvavailed_payslips set amnt4=amnt4+" & mLvWPDay * rsEmployee!daily_rate & " where percode=" & tdbPayrollPeriod.BoundText & " and employeecode=" & rsEmployee!employeecode & ",)"
                    End If
                    
                    ConMain.Execute "insert into lvavailed (percode,employeecode,leavetypescode,lvlimit,lvwpdays,lvamnt) values (" & _
                                    tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & "," & CStr(!leavetypescode) & "," & mLvLimit & "," & mLvWPDay & "," & mLvWPDay * rsEmployee!daily_rate & ")"
                    
                    mLvWPDays = mLvWPDays + mLvWPDay
                    
                  Else
                    mLvWoPDays = mLvWoPDays + ((CDbl(!lvhrs) - CDbl(!lvlimithrs)) / rsParmtr!leavehrsperday)
                  End If
                  
                  CTR = CTR + 1
                  .MoveNext
                  DoEvents
                Loop
              Else
              
                  NetOpen rsTmp, "SELECT * FROM lvlimit WHERE employeecode = " & rsEmployee!employeecode & " AND payyear = '" & tdbPayrollPeriod.Columns("payyear").Text & "' and leavetypescode=1 " 'for SIL only
                  If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    If CDbl(rsTmp!lvlimit) > 0 Then
                      ConMain.Execute "insert into lvavailed_payslips (percode,employeecode,desc1,limit1,days1,amnt1) values (" & tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & ",'SIL'," & CDbl(rsTmp!lvlimit) & ",null,null)"
                    End If
                  End If
                
              End If
              
              mLvWPAmnt = mLvWPDays * rsEmployee!daily_rate
              If rsEmployee!FixedEarnings > 0 Then
                  mFixedEarnings = mFixedEarnings + ((rsEmployee!FixedEarnings / mTtlWrkDays) * mLvWPDays)
              End If
              
            End With
            
            NetOpen rsEarnings, "select x2.otherearningsname, x1.amount from earnings x1 " & _
                                "left outer join otherearnings x2 on x1.otherearningscode = x2.otherearningscode " & _
                                "where x1.percode = '" & tdbPayrollPeriod.BoundText & "' and x1.employeecode = '" & rsEmployee!employeecode & "'"

            CTR = 0
            
            With rsEarnings
                If .RecordCount > 0 Then
                    .MoveFirst
                    CTR = CTR + 1
                    Do While Not .EOF
                        If CTR = 1 Then
                            mFEdesc1 = !OtherEarningsname
                            mFEamnt1 = !amount
                        ElseIf CTR = 2 Then
                            mFEdesc2 = !OtherEarningsname
                            mFEamnt2 = !amount
                        ElseIf CTR = 3 Then
                            mFEdesc3 = !OtherEarningsname
                            mFEamnt3 = !amount
                        ElseIf CTR = 4 Then
                            mFEdesc4 = !OtherEarningsname
                            mFEamnt4 = !amount
                        ElseIf CTR >= 5 Then
                            mFEdesc5 = "Others"
                            mFEamnt5 = mFEamnt5 + !amount
                        End If
                        mEarnings = mEarnings + !amount
                        CTR = CTR + 1
                        .MoveNext
                        DoEvents
                    Loop
                End If
            End With
            
            mRegAmnt = Format(mRegAmnt, "#,##0.00")
            mAbsAmnt = Format(mAbsAmnt, "#,##0.00")
            mLateAmnt = Format(mLateAmnt, "#,##0.00")
            mUtAmnt = Format(mUtAmnt, "#,##0.00")
            mLegAmnt = Format(mLegAmnt, "#,##0.00")
            mSpcAmnt = Format(mSpcAmnt, "#,##0.00")
            mOtRegAmnt = Format(mOtRegAmnt, "#,##0.00")
            mEarnings = Format(mEarnings, "#,##0.00")
            mNonTaxAllowBasic = Format(mNonTaxAllowBasic, "#,##0.00")
            mNonTaxAllowNet = Format(mNonTaxAllowNet, "#,##0.00")
            mLvWPAmnt = Format(mLvWPAmnt, "#,##0.00")
            
            mBasic = mRegAmnt
            
            mGross = (mBasic + mOtRegAmnt + mSpcAmnt + mLegAmnt + mRstAmnt + mFixedEarnings + mCola + mNonTaxAllowNet + mEarnings + mLvWPAmnt + mTtlMealAllow + mNtDfAmnt)
            
            NetOpen rsDeductions, "select x2.otherdeductionsname,x1.amount from deductions x1 " & _
                                    "Left outer join otherdeductions x2 on x1.otherdeductionscode = x2.otherdeductionscode " & _
                                    "where x1.percode = '" & tdbPayrollPeriod.BoundText & "' and x1.employeecode = '" & rsEmployee!employeecode & "'"
            
            CTR = 0
            
            With rsDeductions
                If mFixedDed > 0 Then
                    CTR = CTR + 1
                    mODdesc1 = "Meals"
                    mODamnt1 = mFixedDed
                End If
                If .RecordCount > 0 Then
                    .MoveFirst
                    CTR = CTR + 1
                    Do While Not .EOF
                        If CTR = 1 Then
                            mODdesc1 = !otherdeductionsname
                            mODamnt1 = !amount
                        ElseIf CTR = 2 Then
                            mODdesc2 = !otherdeductionsname
                            mODamnt2 = !amount
                        ElseIf CTR = 3 Then
                            mODdesc3 = !otherdeductionsname
                            mODamnt3 = !amount
                        ElseIf CTR = 4 Then
                            mODdesc4 = !otherdeductionsname
                            mODamnt4 = !amount
                        ElseIf CTR = 5 Then
                            mODdesc5 = !otherdeductionsname
                            mODamnt5 = !amount
                        ElseIf CTR = 6 Then
                            mODdesc6 = !otherdeductionsname
                            mODamnt6 = !amount
                        ElseIf CTR >= 7 Then
                            mODdesc7 = "Others"
                            mODamnt7 = mODamnt7 + !amount
                        End If
                        mDeductions = mDeductions + !amount
                        CTR = CTR + 1
                        .MoveNext
                        DoEvents
                    Loop
                End If
            End With
                
            NetOpen rsLoanDed, "select x2.loantypesname, x1.amtded amount,x1.balance,x1.noofpay  from loanded x1 " & _
                                "left outer join loantypes x2  on x1.loantypescode = x2.loantypescode where x1.fnlz <> 'Y' and " & _
                                "x1.percode = '" & tdbPayrollPeriod.BoundText & "' and x1.employeecode = " & rsEmployee!employeecode & ""
                                
            CTR = 0
            
            With rsLoanDed
                If .RecordCount > 0 Then
                    .MoveFirst
                    CTR = CTR + 1
                    Do While Not .EOF
                        If CTR = 1 Then
                            mLDdesc1 = !loantypesname
                            mLDamnt1 = !amount
                            mLBal1 = !balance
                            mLNoOfPay1 = !noofpay
                        ElseIf CTR = 2 Then
                            mLDdesc2 = !loantypesname
                            mLDamnt2 = !amount
                            mLBal2 = !balance
                            mLNoOfPay2 = !noofpay
                        ElseIf CTR = 3 Then
                            mLDdesc3 = !loantypesname
                            mLDamnt3 = !amount
                            mLBal3 = !balance
                            mLNoOfPay3 = !noofpay
                        ElseIf CTR = 4 Then
                            mLDdesc4 = !loantypesname
                            mLDamnt4 = !amount
                            mLBal4 = !balance
                            mLNoOfPay4 = !noofpay
                        ElseIf CTR = 5 Then
                            mLDdesc5 = !loantypesname
                            mLDamnt5 = !amount
                            mLBal5 = !balance
                            mLNoOfPay5 = !noofpay
                        ElseIf CTR = 6 Then
                            mLDdesc6 = !loantypesname
                            mLDamnt6 = !amount
                            mLBal6 = !balance
                            mLNoOfPay6 = !noofpay
                        ElseIf CTR = 7 Then
                            mLDdesc7 = !loantypesname
                            mLDamnt7 = !amount
                            mLBal7 = !balance
                            mLNoOfPay7 = !noofpay
                        ElseIf CTR >= 8 Then
                            mLDdesc8 = "Others"
                            mLDamnt8 = mLDamnt8 + !amount
                        End If
                        mLoanded = mLoanded + !amount
                        CTR = CTR + 1
                        .MoveNext
                        DoEvents
                    Loop
                End If
            End With
            
            If ((rsEmployee!ratetypecode = 4 And rsPayPeriod!sssmonthly = "Y") Or (rsEmployee!ratetypecode = 1 And rsPayPeriod!sssdaily = "Y")) And mBasic > 0 Then
                
                mSSSBasic = (mBasic + mOtRegAmnt + mSpcAmnt + mLegAmnt + mRstAmnt + mFixedEarnings + _
                             mCola + mNonTaxAllowNet + mLvWPAmnt + mTtlMealAllow + mNtDfAmnt) - _
                            (mAbsAmnt + mUtAmnt + mLateAmnt)
                
                If rsEmployee!sssauto = 0 Then
                    ttlEE = rsEmployee!sssamt
                    ttlER = rsEmployee!SssEr
                Else
                    CompSss rsEmployee!employeecode, tdbPayrollPeriod.BoundText, _
                            tdbPayrollPeriod.Columns("payyear").Text, tdbPayrollPeriod.Columns("paymonth").Text, _
                            mSSSBasic, regER, regEE, ecER, ecEE, mpfER, mpfEE, ttlEE, ttlER
                End If
                
                ConMain.Execute "delete from payroll_sss_contributions " & _
                                "where percode =" & tdbPayrollPeriod.BoundText & " and " & _
                                "employeecode = " & rsEmployee!employeecode
                                
                ConMain.Execute "insert into payroll_sss_contributions " & _
                                "(percode,employeecode,divisioncode,costcentercode,payyear,paymonth,basic_amt," & _
                                "reg_er,reg_ee,ec_er,ec_ee,mpf_er,mpf_ee,ttl_er,ttl_ee) " & _
                                "values (" & tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & "," & _
                                rsEmployee!divisioncode & "," & rsEmployee!costcentercode & "," & _
                                tdbPayrollPeriod.Columns("payyear").Text & "," & _
                                "'" & tdbPayrollPeriod.Columns("paymonth").Text & "'," & mSSSBasic & ", " & _
                                regER & "," & regEE & "," & ecER & "," & ecEE & "," & mpfER & "," & mpfEE & "," & _
                                ttlER & " ," & ttlEE & ")"
                                
            End If
            
            If ((rsEmployee!ratetypecode = 4 And rsPayPeriod!phmonthly = "Y") Or (rsEmployee!ratetypecode = 1 And rsPayPeriod!phdaily = "Y")) And rsEmployee!monthly_rate > 0 Then
                If rsEmployee!philhauto = 0 Then
                    mPhilAmnt = rsEmployee!PhilHAmt
                    mPhilEr = rsEmployee!philher
                Else
                
                    philBasic = (mRegAmnt - mAbsAmnt - mUtAmnt) + mLegAmnt + mLvWPAmnt + mRstAmnt
                    CompPhi philBasic, mPhilAmnt, mPhilEr
                    
                    '## PhilHealth Contribution Computaion as of 2024-05-25
                    'CompPhi rsEmployee!monthly_rate, mPhilAmnt, mPhilEr
                End If
            End If
            
            If ((rsEmployee!ratetypecode = 4 And rsPayPeriod!hdmfdaily = "Y") Or (rsEmployee!ratetypecode = 1 And rsPayPeriod!hdmfmonthly = "Y")) Then
                If rsEmployee!hdmfauto = 0 Then
                    mHdmfAmnt = rsEmployee!HdmfAmt
                    mHdmfEr = rsEmployee!HDMFEr
                End If
            End If
            
            mTaxableInc = mBasic - (mAbsAmnt + mUtAmnt + mLateAmnt + ttlEE + mPhilAmnt + mHdmfAmnt)
            
            If ((rsEmployee!ratetypecode = 4 And rsPayPeriod!taxmonthly = "Y") Or (rsEmployee!ratetypecode = 1 And rsPayPeriod!taxdaily = "Y")) And mTaxableInc > 0 Then
                If Trim(rsEmployee!tinno) <> "" Then
                  If rsEmployee!taxauto = 0 Then
                      mTaxAmnt = rsEmployee!taxamt
                  Else
                    If Not IsNull(rsEmployee!wtcode) Then
                      CompTax mTaxableInc, rsEmployee!wtcode, mTaxAmnt
                    End If
                  End If
                End If
            End If
            
            ttlEE = Format(ttlEE, "#,##0.00")
            mPhilAmnt = Format(mPhilAmnt, "#,##0.00")
            mHdmfAmnt = Format(mHdmfAmnt, "#,##0.00")
            mTaxAmnt = Format(mTaxAmnt, "#,##0.00")
            mDeductions = Format(mDeductions, "#,##0.00")
            mLoanded = Format(mLoanded, "#,##0.00")
            mTaxableInc = Format(mTaxableInc, "#,##0.00")
            
            mNet = mGross - (ttlEE + mPhilAmnt + mHdmfAmnt + mTaxAmnt + mDeductions + mLoanded + mAbsAmnt + mLateAmnt + mUtAmnt + mFixedDed)
            
            ConMain.Execute "delete from payroll where employeecode = '" & rsEmployee!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "'"
            
            ConMain.Execute "insert into payroll(percode,branchcode,divisioncode,costcentercode,employeecode,ratetypecode,empstatcode,regamnt,dayswrk, " & _
                        "absdays,absamnt, regwrkhrs,latehrs,lateamnt,uthrs,utamnt,legdays,leghrs,legamnt,spcdays,spchrs,spcamnt,rstdays,rsthrs,rstamnt, " & _
                        "otreghrs,otregamnt,ntdfhrs,ntdfamnt,daily_rate,hourly_rate,fixeddeductions,nontaxallow_basic,nontaxallow_net, " & _
                        "earnings,gross,sssamnt,ssser,ec,philamnt,philer,hdmfamnt,monthly_rate,bankcode,bankacctno,saltobank,regular, " & _
                        "hdmfer,taxamnt, deductions,net,basic,lvwpdays,lvwpamnt,lvwopdays,lvwopamnt,loanded,payyear,paymonth,mealallow,fixedearnings,cola," & _
                        "vllimit,vldays,vlamnt,sllimit,sldays,slamnt,wtcode,tinno,sssno,philhno, " & _
                        "fedesc1,fedesc2,fedesc3,fedesc4,fedesc5," & _
                        "feamnt1,feamnt2,feamnt3,feamnt4,feamnt5," & _
                        "oddesc1,oddesc2,oddesc3,oddesc4,oddesc5,oddesc6,oddesc7," & _
                        "odamnt1,odamnt2,odamnt3,odamnt4,odamnt5,odamnt6,odamnt7,lbal1,lbal2,lbal3,lbal4,lbal5,lbal6,lbal7," & _
                        "lddesc1,lddesc2,lddesc3,lddesc4,lddesc5,lddesc6,lddesc7,lddesc8,lnoofpay1,lnoofpay2,lnoofpay3,lnoofpay4,lnoofpay5,lnoofpay6,lnoofpay7," & _
                        "ldamnt1,ldamnt2,ldamnt3,ldamnt4,ldamnt5,ldamnt6,ldamnt7,ldamnt8,fnlz,withsss,withphilh,taxableincome,confidential) values " & _
                        "(" & tdbPayrollPeriod.BoundText & ",'" & rsEmployee!branchcode & "','" & rsEmployee!divisioncode & "','" & rsEmployee!costcentercode & "','" & rsEmployee!employeecode & "','" & rsEmployee!ratetypecode & "','" & rsEmployee!empstatcode & "'," & mRegAmnt & "," & mDaysWrk & ", " & _
                        mAbsDays & "," & mAbsAmnt & "," & mRegWrkHrs & "," & mLatehrs & "," & mLateAmnt & "," & mUtHrs & "," & mUtAmnt & "," & mLegDays & "," & mLegHrs & "," & mLegAmnt & "," & mSpcDays & "," & mSpcHrs & "," & mSpcAmnt & "," & mRstDays & "," & mRstHrs & "," & mRstAmnt & ", " & _
                        mOtRegHrs & ", " & mOtRegAmnt & "," & mNtDfHrs & "," & mNtDfAmnt & "," & rsEmployee!daily_rate & "," & rsEmployee!Hourly_Rate & "," & mFixedDed & "," & mNonTaxAllowBasic & "," & mNonTaxAllowNet & ", " & _
                        mEarnings & "," & mGross & "," & ttlEE & "," & ttlER & ",0," & mPhilAmnt & "," & mPhilEr & "," & mHdmfAmnt & "," & rsEmployee!monthly_rate & "," & IIf(IsNull(rsEmployee!bankcode), "Null", rsEmployee!bankcode) & ",'" & rsEmployee!bankacctno & "','" & rsEmployee!saltobank & "','" & rsEmployee!regular & "', " & _
                        mHdmfEr & "," & mTaxAmnt & "," & mDeductions & "," & mNet & "," & mBasic & "," & mLvWPDays & "," & mLvWPAmnt & "," & mLvWoPDays & "," & mLvWoPAmnt & "," & mLoanded & ",'" & tdbPayrollPeriod.Columns("payyear").Text & "','" & tdbPayrollPeriod.Columns("paymonth").Text & "'," & mTtlMealAllow & "," & mFixedEarnings & "," & mCola & ", " & _
                        0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & IIf(IsNumeric(rsEmployee!wtcode), rsEmployee!wtcode, "Null") & "," & IIf(Not IsNull(rsEmployee!tinno), "'" & rsEmployee!tinno & "'", "Null") & "," & IIf(IsNull(rsEmployee!sssno), "Null", "'" & rsEmployee!sssno & "'") & "," & IIf(IsNull(rsEmployee!philhno), "Null", "'" & rsEmployee!philhno & "'") & ",'" & _
                        mFEdesc1 & "','" & mFEdesc2 & "','" & mFEdesc3 & "','" & mFEdesc4 & "','" & mFEdesc5 & "'," & _
                        mFEamnt1 & "," & mFEamnt2 & "," & mFEamnt3 & "," & mFEamnt4 & "," & mFEamnt5 & ",'" & _
                        Swap(mODdesc1) & "','" & Swap(mODdesc2) & "','" & Swap(mODdesc3) & "','" & Swap(mODdesc4) & "','" & Swap(mODdesc5) & "','" & Swap(mODdesc6) & "','" & Swap(mODdesc7) & "'," & _
                        mODamnt1 & "," & mODamnt2 & "," & mODamnt3 & "," & mODamnt4 & "," & mODamnt5 & "," & mODamnt6 & "," & mODamnt7 & "," & mLBal1 & "," & mLBal2 & "," & mLBal3 & "," & mLBal4 & "," & mLBal5 & "," & mLBal6 & "," & mLBal7 & ",'" & _
                        mLDdesc1 & "','" & mLDdesc2 & "','" & mLDdesc3 & "','" & mLDdesc4 & "','" & mLDdesc5 & "','" & mLDdesc6 & "','" & mLDdesc7 & "','" & mLDdesc8 & "'," & mLNoOfPay1 & "," & mLNoOfPay2 & "," & mLNoOfPay3 & "," & mLNoOfPay4 & "," & mLNoOfPay5 & "," & mLNoOfPay6 & "," & mLNoOfPay7 & "," & _
                        mLDamnt1 & "," & mLDamnt2 & "," & mLDamnt3 & "," & mLDamnt4 & "," & mLDamnt5 & "," & mLDamnt6 & "," & mLDamnt7 & "," & mLDamnt8 & ",'Y','" & rsEmployee!withsss & "','" & rsEmployee!withphilh & "'," & mTaxableInc & ",'" & rsEmployee!confidential & "')"
            
            
            rsEmployee.MoveNext
            
            DoEvents
            
        Loop
        
        NetOpen rsChk, "select count(percode) ctr from payroll where percode = " & tdbPayrollPeriod.BoundText & " " 'and fnlz = 'Y'"
        
        mPRCtr = rsChk!CTR
        
        NetOpen rsChk, "select count(percode) ctr from dtr where percode = " & tdbPayrollPeriod.BoundText & ""
        
        mDTRCtr = rsChk!CTR
        
        If mPRCtr = mDTRCtr Then
            ConMain.Execute "update payrollperiod set genpay = 'Y' where percode = " & tdbPayrollPeriod.BoundText & ""
        Else
            ConMain.Execute "update payrollperiod set genpay = 'N' where percode = " & tdbPayrollPeriod.BoundText & ""
        End If
        
        ConMain.CommitTrans
        fra1.Enabled = True
        cmdGenerate.Enabled = True
        Me.MousePointer = vbDefault
        
        MsgBox "Process completed!", vbInformation + vbOKOnly
        
        pb1.Value = 0
        pb2.Value = 0
        
    Else
        MsgBox "No record found.", vbExclamation + vbOKOnly
    End If

End Sub

Private Sub CompNonTaxAllow(ByVal rsDtrTmp As ADODB.Recordset, ByVal mTtlWrkHrs As Double, ByVal mDaysWrk As Double, ByVal mAbsDays As Double, ByVal mLatehrs As Double, ByVal mUtHrs As Double, ByRef mBasicTtl As Double, ByRef mNetTtl As Double)
  
  Dim rsTmp                 As ADODB.Recordset
  Dim mBasicAmt             As Double
  Dim mNetAmt               As Double
  Dim mNTA_Data(0 To 3)     As NTA_Data
  Dim mCtr                  As Integer
  
  mCtr = 0
  ConMain.Execute "delete from payroll_nontaxallow where percode = " & tdbPayrollPeriod.BoundText & " and employeecode = " & rsEmployee!employeecode & ""
  ConMain.Execute "delete from payroll_nontaxallow_payslips where percode = " & tdbPayrollPeriod.BoundText & " and employeecode = " & rsEmployee!employeecode & ""
  
  NetOpen rsTmp, "select x1.*,x2.nontaxallow_description from employee_nontaxallow x1 " & _
                 "left outer join nontaxallow x2 on x1.nontaxallow_id=x2.nontaxallow_id " & _
                 "where x1.employeecode = " & rsEmployee!employeecode & ""
  If rsTmp.RecordCount > 0 Then
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
      mNetAmt = 0
      If rsEmployee!ratetypecode = 4 Then 'Monthly Rate Employee
        mBasicAmt = rsTmp!nontaxallow_amt / 2
                 '********Hours Worked*********     *****absent hrs + late + undertime ******
        mNetAmt = (rsTmp!nontaxallow_amt / mTtlWrkHrs) * ((mTtlWrkHrs / 2) - ((mAbsDays * 8) + mLatehrs + mUtHrs))
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDtrTmp!legdays)
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDtrTmp!spcdays)
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * rsDtrTmp!restdays)
      ElseIf rsEmployee!ratetypecode = 1 Then 'Daily Rate Employee
        mBasicAmt = ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * (mDaysWrk * 8))
                 '*********hourly rate**********    **hrs worked**   *********hourly rate**********   *no. of hrs tardy*
        mNetAmt = ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * (mDaysWrk * 8)) - ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * (mLatehrs + mUtHrs))
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * Format(rsParmtr!legholprct / 100, "#,##0.00") * rsDtrTmp!legdays)
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * Format(rsParmtr!spcholprct / 100, "#,##0.00") * rsDtrTmp!spcdays)
        mNetAmt = mNetAmt + ((rsTmp!nontaxallow_amt / mTtlWrkHrs) * rsDtrTmp!restdays)
      End If
      ConMain.Execute "insert into payroll_nontaxallow (percode,employeecode,nontaxallow_id,nontaxallow_basic,nontaxallow_net) values (" & _
                       tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & "," & rsTmp!nontaxallow_id & "," & mBasicAmt & "," & mNetAmt & ")"
      'mBasicTtl = mBasicTtl + mBasicAmt  ***2017
      mBasicTtl = mBasicTtl + (rsTmp!nontaxallow_amt / 2)
      mNetTtl = mNetTtl + mNetAmt
      If mCtr < 3 Then
        mNTA_Data(mCtr).mNTA_Desc = rsTmp!nontaxallow_description
        mNTA_Data(mCtr).mNTA_Amt = mNetAmt
      Else
        mNTA_Data(3).mNTA_Amt = mNTA_Data(3).mNTA_Amt + mNetAmt
      End If
      mCtr = mCtr + 1
      rsTmp.MoveNext
    Loop
    ConMain.Execute "insert into payroll_nontaxallow_payslips (percode,employeecode,nta_desc1,nta_amnt1," & _
                    "nta_desc2,nta_amnt2,nta_desc3,nta_amnt3,nta_amnt4) values (" & _
                    tdbPayrollPeriod.BoundText & "," & rsEmployee!employeecode & ",'" & mNTA_Data(0).mNTA_Desc & "'," & mNTA_Data(0).mNTA_Amt & ", '" & _
                     mNTA_Data(1).mNTA_Desc & "'," & mNTA_Data(1).mNTA_Amt & ",'" & mNTA_Data(2).mNTA_Desc & "'," & mNTA_Data(2).mNTA_Amt & "," & mNTA_Data(3).mNTA_Amt & ")"
  End If
  
        
End Sub

Private Sub CompSss(ByVal empID As Integer, ByVal percode As Integer, ByVal payYear As Integer, ByVal payMonth As String, _
                    ByVal basicAmt As Double, ByRef regER As Double, ByRef regEE As Double, _
                    ByRef ecER As Double, ByRef ecEE As Double, ByRef mpfER As Double, ByRef mpfEE As Double, _
                    ByRef ttlEE As Double, ByRef ttlER As Double)

    Dim rsTmp   As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    
    Dim ttlBasic      As Double
    Dim prevRegER     As Double
    Dim prevRegEE     As Double
    Dim prevEcER      As Double
    Dim prevEcEE      As Double
    Dim prevMpfER     As Double
    Dim prevMpfEE     As Double
    
    NetOpen rsTmp, "SELECT IFNULL(SUM(basic_amt),0) basic_amt,IFNULL(SUM(reg_er),0) reg_er,IFNULL(SUM(reg_ee),0) reg_ee,IFNULL(SUM(ec_er),0) ec_er," & _
                   "IFNULL(SUM(ec_ee),0) ec_ee,IFNULL(SUM(mpf_er),0) mpf_er,IFNULL(SUM(mpf_ee),0) mpf_ee, " & _
                   "IFNULL(SUM(ttl_er),0) ttl_er,IFNULL(SUM(ttl_ee),0) ttl_ee " & _
                   "From payroll_sss_contributions " & _
                   "WHERE employeecode =" & empID & " AND payyear=" & payYear & _
                   " AND paymonth='" & payMonth & "' and fnlz='Y'"
    
    ttlBasic = rsTmp!basic_amt + basicAmt
    prevRegER = rsTmp!reg_er
    prevRegEE = rsTmp!reg_ee
    prevEcER = rsTmp!ec_er
    prevEcEE = rsTmp!ec_ee
    prevMpfER = rsTmp!mpf_er
    prevMpfEE = rsTmp!mpf_ee
    
    NetOpen rsTmp, "select * from sss order by contribution_id"
    
    If rsTmp.RecordCount > 0 Then
      With rsTmp
        .MoveFirst
        Do While Not .EOF
            If (ttlBasic >= .Fields("fromamount")) And (ttlBasic <= .Fields("toamount")) Then
              regER = rsTmp!reg_er - prevRegER
              regEE = rsTmp!reg_ee - prevRegEE
              ecER = rsTmp!ec_er - prevEcER
              ecEE = rsTmp!ec_ee - prevEcEE
              mpfER = rsTmp!mpf_er - prevMpfER
              mpfEE = rsTmp!mpf_ee - prevMpfEE
              ttlER = regER + ecER + mpfER
              ttlEE = regEE + ecEE + mpfEE
              Exit Do
            End If
            .MoveNext
        Loop
      End With
    End If
    
End Sub

Private Sub CompSss20210110(ByVal par1 As Double, ByRef ttlEE As Double, ByRef ttlER As Double, ByRef contID As Integer)

    Dim rsTmp   As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    
    NetOpen rsTmp, "select * from sss order by contribution_id"
    
    If rsTmp.RecordCount > 0 Then
      With rsTmp
        .MoveFirst
        Do While Not .EOF
        
            If (par1 >= .Fields("fromamount")) And (par1 <= .Fields("toamount")) Then
              ttlER = .Fields("ttl_er")
              ttlEE = .Fields("ttl_ee")
              contID = .Fields("contribution_id")
              Exit Do
            End If
            
            .MoveNext
            
        Loop
      End With
    End If
    
End Sub

'## PhilHealth Contribution Computaion as of 2024-05-25
'Private Sub CompPhi(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double)
'
'  Dim rsTmp   As ADODB.Recordset
'  Set rsTmp = New ADODB.Recordset
'
'  NetOpen rsTmp, "select * from ph"
'
'  If rsTmp.RecordCount > 0 Then
'    If par1 <= rsTmp!income_floor Then
'      mAmount = Format(rsTmp!income_floor * rsTmp!rate_prct / 2, "#,##0.00")
'    ElseIf par1 >= rsTmp!income_ceiling Then
'      mAmount = Format(rsTmp!income_ceiling * rsTmp!rate_prct / 2, "#,##0.00")
'    Else
'      mAmount = Format(par1 * rsTmp!rate_prct / 2, "#,##0.00")
'    End If
'    mEr = mAmount
'  End If
'
'End Sub

Private Sub CompPhi(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double)
  
  Dim rsTmp   As ADODB.Recordset
  Set rsTmp = New ADODB.Recordset
    
  NetOpen rsTmp, "select * from ph"
  
  If rsTmp.RecordCount > 0 Then
    If par1 <= rsTmp!income_floor Then
      mAmount = Format(rsTmp!income_floor * rsTmp!rate_prct, "#,##0.00")
    ElseIf par1 >= rsTmp!income_ceiling Then
      mAmount = Format(rsTmp!income_ceiling * rsTmp!rate_prct, "#,##0.00")
    Else
      mAmount = Format(par1 * rsTmp!rate_prct, "#,##0.00")
    End If
    mEr = mAmount
  End If
  
End Sub


Private Sub CompPhi_2019(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double)
    If par1 < 10000 Then par1 = 10000
    mAmount = par1 * 0.0275 / 2
    mEr = par1 * 0.0275 / 2
End Sub

Private Sub CompPhi_2017(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double)
        
    Dim rsTmp   As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    NetOpen rsTmp, "select * from ph order by fromamount"
    If rsTmp.RecordCount > 0 Then
        With rsTmp
          .MoveFirst
          Do While Not .EOF
          
            If (par1 >= .Fields("fromamount")) And (par1 <= .Fields("toamount")) Then
              If .Fields("ERAmount") > 0 Then
                mAmount = .Fields("EEAmount")
                mEr = .Fields("ERAmount")
              Else
                mAmount = Format(par1 * .Fields("EEPercent"), "#,##0.00")
                mEr = Format(par1 * .Fields("ERPercent"), "#,##0.00")
              End If
              
              Exit Do
            End If
            If .AbsolutePosition = .RecordCount Then
              If .Fields("ERAmount") > 0 Then
                mAmount = .Fields("EEAmount")
                mEr = .Fields("ERAmount")
              Else
                mAmount = Format(par1 * .Fields("EEPercent"), "#,##0.00")
                mEr = Format(par1 * .Fields("ERPercent"), "#,##0.00")
              End If
              Exit Do
            End If
            .MoveNext
          Loop
        End With
    End If
    
End Sub

Private Sub CompTax(ByVal par1 As Double, ByVal mCode As String, ByRef mAmount As Double)

    Dim rsTmp As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    
    NetOpen rsTmp, "select * from wt where wtcode = '" & mCode & "'"
    With rsTmp
        If .RecordCount > 0 Then
          If par1 <= .Fields("b1") Then
            mAmount = 0
          ElseIf (par1 > .Fields("b1")) And (par1 <= .Fields("b2")) Then
            mAmount = 0
          ElseIf (par1 > .Fields("b2")) And (par1 <= .Fields("b3")) Then
            mAmount = Format((par1 - .Fields("b2")) * (.Fields("f2") / 100), "#,##0.00")
          ElseIf (par1 > .Fields("b3")) And (par1 <= .Fields("b4")) Then
            mAmount = Format(((par1 - .Fields("b3")) * (.Fields("f3") / 100)) + .Fields("a3"), "#,##0.00")
          ElseIf (par1 > .Fields("b4")) And (par1 <= .Fields("b5")) Then
            mAmount = Format(((par1 - .Fields("b4")) * (.Fields("f4") / 100)) + .Fields("a4"), "#,##0.00")
          ElseIf (par1 > .Fields("b5")) And (par1 <= .Fields("b6")) Then
            mAmount = Format(((par1 - .Fields("b5")) * (.Fields("f5") / 100)) + .Fields("a5"), "#,##0.00")
          ElseIf (par1 > .Fields("b6")) And (par1 <= .Fields("b7")) Then
            mAmount = Format(((par1 - .Fields("b6")) * (.Fields("f6") / 100)) + .Fields("a6"), "#,##0.00")
          ElseIf (par1 > .Fields("b7")) And (par1 <= .Fields("b8")) Then
            mAmount = Format(((par1 - .Fields("b7")) * (.Fields("f7") / 100)) + .Fields("a7"), "#,##0.00")
          ElseIf (par1 > .Fields("b8")) Then 'And (par1 <= .Fields("b9")) Then
            mAmount = Format(((par1 - .Fields("b8")) * (.Fields("f8") / 100)) + .Fields("a8"), "#,##0.00")
    '      ElseIf (par1 > .Fields("b9")) And (par1 <= .Fields("brckt10")) Then
    '        mamount = ((par1 - .Fields("b9")) * (.Fields("f9") / 100)) + .Fields("a9")
    '      ElseIf (par1 > .Fields("brckt10")) Then
    '        mamount = ((par1 - .Fields("brckt10")) * (.Fields("f10") / 100)) + .Fields("a10")
          End If
        End If
    End With
End Sub

Private Sub tdbEmployee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub tdbPayrollPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
    End If
End Sub



