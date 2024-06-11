VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPGenPay 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Payroll"
   ClientHeight    =   3015
   ClientLeft      =   6330
   ClientTop       =   5265
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPPGenPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   -15
      TabIndex        =   0
      Top             =   -90
      Width           =   6180
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   2715
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         Caption         =   "Proceed"
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
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   105
         Left            =   1980
         TabIndex        =   2
         Top             =   2745
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
         Top             =   2160
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
         Appearance      =   3
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
         _PropDict       =   $"frmPPGenPay.frx":6852
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
         TabIndex        =   4
         Tag             =   "Municipal"
         Top             =   855
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
         Appearance      =   3
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
         _PropDict       =   $"frmPPGenPay.frx":68FC
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
      Begin TrueOleDBList80.TDBCombo tdbCostCenter 
         Height          =   345
         Left            =   1980
         TabIndex        =   5
         Tag             =   "Municipal"
         Top             =   1635
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
         Appearance      =   3
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
         _PropDict       =   $"frmPPGenPay.frx":69A6
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
         TabIndex        =   6
         Tag             =   "Municipal"
         Top             =   1245
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
         Appearance      =   3
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
         _PropDict       =   $"frmPPGenPay.frx":6A50
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
         Top             =   2880
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
         TabIndex        =   13
         Tag             =   "Municipal"
         Top             =   255
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
         _PropDict       =   $"frmPPGenPay.frx":6AFA
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
         ForeColor       =   &H0030A0B8&
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   2235
         Width           =   1710
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
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   345
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   1665
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
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   1305
         Width           =   1455
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
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Index           =   2
         Left            =   465
         TabIndex        =   8
         Top             =   930
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0030A0B8&
         X1              =   120
         X2              =   6000
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0030A0B8&
         X1              =   120
         X2              =   6000
         Y1              =   2040
         Y2              =   2040
      End
   End
End
Attribute VB_Name = "frmPPGenPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsEmployee      As ADODB.Recordset

Private Sub Form_Load()
    
    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    SendMessage pb2.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb2.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    bind_tdb ConMain, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto,payfreqcode,payyear,paymonth from payrollperiod", "description", "percode"
    
    bind_tdb ConMain, tdbBranch, "select branchcode, branch from branch order by branch", "branch", "branchcode"
  
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
End Sub

Private Sub tdbBranch_ItemChange()

  bind_tdb ConMain, tdbDivision, "select divisioncode,division from division " & _
            "where branchcode = " & tdbBranch.BoundText & " order by division", "division", "divisioncode"
  
  If tdbDivision.ApproxCount > 0 Then
    tdbDivision.BoundText = tdbDivision.Columns(0).Text
  End If
  
End Sub

Private Sub tdbBranch_Keypress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbBranch, tdbBranch.RowSource, tdbBranch.Text
    tdbBranch_ItemChange
  End If
  
End Sub

Private Sub tdbDivision_Itemchange()

    tdbCostCenter.BoundText = ""
  
  bind_tdb ConMain, tdbCostCenter, "select costcentercode,costcenter from costcenter " & _
            "where branchcode = '" & tdbBranch.BoundText & "' order by costcenter", "costcenter", "costcentercode"
  
End Sub

Private Sub tdbDivision_Keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchList KeyAscii, tdbDivision, tdbDivision.RowSource, tdbDivision.Text
      tdbDivision_Itemchange
    End If
    
End Sub

Private Sub tdbcostcenter_ItemChange()

    tdbEmployee.BoundText = ""
    bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname, ', ', firstname,' ',middlename) fullname from employee " & _
                "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and costcentercode = '" & tdbCostCenter.BoundText & "' " & _
                "order by concat(lastname, ', ', firstname,' ',middlename) ", "fullname", "employeecode"
    
End Sub

Private Sub tdbcostcenter_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchList KeyAscii, tdbCostCenter, tdbCostCenter.RowSource, tdbCostCenter.Text
      tdbcostcenter_ItemChange
    End If
    
End Sub

Private Sub cmdGenerate_Click()

    Dim rsDtrEmp        As ADODB.Recordset
    Dim rsDtrEmpTmp     As ADODB.Recordset
    Dim rsOT            As ADODB.Recordset
    Dim rsDts           As ADODB.Recordset
    Dim rsParmtr        As ADODB.Recordset
    Dim rsLeave         As ADODB.Recordset
    Dim rsEarnings      As ADODB.Recordset
    Dim rsDeductions    As ADODB.Recordset
    Dim rsLoanDed       As ADODB.Recordset
    
    Dim mBasic          As Double
    Dim mGross          As Double
    Dim mNet            As Double
    
    Dim mRstDays        As Double
    Dim mRstHrs         As Double
    Dim mRstAmnt        As Double
    
    Dim mLegDays        As Double
    Dim mLegHrs         As Double
    Dim mLegAmnt        As Double
    
    Dim mSpcDays        As Double
    Dim mSpcHrs         As Double
    Dim mSpcAmnt        As Double
    
    Dim mRstLegDays     As Double
    Dim mRstLegHrs      As Double
    Dim mRstLegAmnt     As Double
    
    Dim mRstSpcDays     As Double
    Dim mRstSpcHrs      As Double
    Dim mRstSpcAmnt     As Double
    
    Dim mLvWPDays       As Double
    Dim mLvWPAmnt       As Double
    Dim mLvWoPDays      As Double
    Dim mLvWoPAmnt      As Double
    
    Dim mHrlyRate       As Double
    Dim mDailyRate      As Double
    Dim mRegAmnt        As Double
    Dim mPORate         As Double
    
    Dim mRegWrkHrs      As Double
    Dim mDaysWrk        As Double
    Dim mAbsDays        As Double
    Dim mLatehrs        As Double
    Dim mUtHrs          As Double
    
    Dim mAbsAmnt        As Double
    Dim mLateAmnt       As Double
    Dim mUtAmnt         As Double
    
    Dim mOtRegHrs       As Double
    Dim mOtRstHrs       As Double
    Dim mOtSpcHrs       As Double
    Dim mOtLegHrs       As Double
        
    Dim mOtRegAmnt      As Double
    Dim mOtRstAmnt      As Double
    Dim mOtSpcAmnt      As Double
    Dim mOtLegAmnt      As Double
    
    Dim mOtNpRegHrs     As Double
    Dim mOtNpRstHrs     As Double
    Dim mOtNpSpcHrs     As Double
    Dim mOtNpLegHrs     As Double
    
    Dim mOtNpRegAmnt    As Double
    Dim mOtNpRstAmnt    As Double
    Dim mOtNpSpcAmnt    As Double
    Dim mOtNpLegAmnt    As Double
    
    Dim mNpRegHrs       As Double
    Dim mNpRstHrs       As Double
    Dim mNpSpcHrs       As Double
    Dim mNpLegHrs       As Double
    
    Dim mNpRegAmnt      As Double
    Dim mNpRstAmnt      As Double
    Dim mNpSpcAmnt      As Double
    Dim mNpLegAmnt      As Double
    
    Dim mPORegHrs       As Double
    Dim mPORstHrs       As Double
    Dim mPOSpcHrs       As Double
    Dim mPOLegHrs       As Double
      
    Dim mPORegAmnt      As Double
    Dim mPORstAmnt      As Double
    Dim mPOSpcAmnt      As Double
    Dim mPOLegAmnt      As Double
    
    Dim mPONpRegHrs     As Double
    Dim mPONpRstHrs     As Double
    Dim mPONpSpcHrs     As Double
    Dim mPONpLegHrs     As Double
    
    Dim mPONpRegAmnt    As Double
    Dim mPONpRstAmnt    As Double
    Dim mPONpSpcAmnt    As Double
    Dim mPONpLegAmnt    As Double
    
    Dim mEarnings       As Double
    Dim mDeductions     As Double
    Dim mLoanded        As Double
    
    Dim mMealHrlyRate   As Double
    Dim mTtlMealAllow   As Double
    
    Dim mFEHrlyRate     As Double
    Dim mFixedEarnings  As Double
    
    Dim mSssAmnt        As Double
    Dim mSSSEr          As Double
    Dim mSSSEc          As Double
    
    Dim mPhilAmnt       As Double
    Dim mPhilEr         As Double
    
    Dim mHdmfAmnt       As Double
    Dim mHdmfEr         As Double
    Dim mTaxAmnt        As Double
    
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
    
    If MsgBox("Confirm generate payroll.", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
        If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
            If Trim(tdbCostCenter.Text) <> "" And Not IsNull(tdbCostCenter.SelectedItem) And tdbCostCenter.ApproxCount > 0 Then
                If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
                
                    NetOpen rsEmployee, "select * from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' and  " & _
                                            "costcentercode = '" & tdbCostCenter.BoundText & "' and employeecode = '" & tdbEmployee.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                                            
                Else
                
                    NetOpen rsEmployee, "select * from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' and  " & _
                                            "costcentercode = '" & tdbCostCenter.BoundText & "' and " & _
                                            "payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                                            
                End If
            Else
                
                NetOpen rsEmployee, "select * from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                                            
            End If
        Else
        
            NetOpen rsEmployee, "select * from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
                                            
        End If
    Else
    
        NetOpen rsEmployee, "select * from employee where payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
        
    End If
   
    If rsEmployee.RecordCount > 0 Then
    
        fra1.Enabled = False
        Me.MousePointer = vbHourglass
        cmdGenerate.Enabled = False
        
        pb1.Max = rsEmployee.RecordCount
        pb1.Value = 0
        
        rsEmployee.MoveFirst
        
        mPORate = rsParmtr!pulloutrate / rsParmtr!hrsperday
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        
        Do While Not rsEmployee.EOF
        
            pb1.Value = pb1.Value + 1
            
            mHrlyRate = 0
            mDailyRate = 0
            mPORate = 0
            mRegAmnt = 0
            
            mBasic = 0
            mGross = 0
            mNet = 0
            
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
            mRstLegDays = 0
            mRstLegHrs = 0
            mRstLegAmnt = 0
            mRstSpcDays = 0
            mRstSpcHrs = 0
            mRstSpcAmnt = 0
            
            mRegWrkHrs = 0
            mDaysWrk = 0
            mAbsDays = 0
            mLatehrs = 0
            mUtHrs = 0
            
            mAbsAmnt = 0
            mLateAmnt = 0
            mUtAmnt = 0
            
            mOtRegHrs = 0
            mOtRstHrs = 0
            mOtSpcHrs = 0
            mOtLegHrs = 0
            
            mOtRegAmnt = 0
            mOtRstAmnt = 0
            mOtSpcAmnt = 0
            mOtLegAmnt = 0
            
            mOtNpRegHrs = 0
            mOtNpRstHrs = 0
            mOtNpSpcHrs = 0
            mOtNpLegHrs = 0
            
            mOtNpRegAmnt = 0
            mOtNpRstAmnt = 0
            mOtNpSpcAmnt = 0
            mOtNpLegAmnt = 0
            
            mNpRegHrs = 0
            mNpRstHrs = 0
            mNpSpcHrs = 0
            mNpLegHrs = 0
                    
            mNpRegAmnt = 0
            mNpRstAmnt = 0
            mNpSpcAmnt = 0
            mNpLegAmnt = 0
            
            mPORstHrs = 0
            mPORegHrs = 0
            mPOSpcHrs = 0
            mPOLegHrs = 0
            
            mPORstAmnt = 0
            mPORegAmnt = 0
            mPOSpcAmnt = 0
            mPOLegAmnt = 0
            
            mPONpRegHrs = 0
            mPONpRstHrs = 0
            mPONpSpcHrs = 0
            mPONpLegHrs = 0
                    
            mPONpRegAmnt = 0
            mPONpRstAmnt = 0
            mPONpSpcAmnt = 0
            mPONpLegAmnt = 0
            
            mEarnings = 0
            mDeductions = 0
            mLoanded = 0
            
            mMealHrlyRate = 0
            mTtlMealAllow = 0
            
            mFEHrlyRate = 0
            mFixedEarnings = 0
            
            mSssAmnt = 0
            mSSSEr = 0
            mSSSEc = 0
            
            mPhilAmnt = 0
            mPhilEr = 0
            
            mHdmfAmnt = 0
            mHdmfEr = 0
            mTaxAmnt = 0
            
            If rsEmployee!ratetypecode = "0000004" Then  'monthly rate employees
              mDailyRate = Format((rsEmployee!payrate * 12) / 312, "#,##0.00")
              mRegAmnt = Format(rsEmployee!payrate / 2, "#,##0.00")
            ElseIf rsEmployee!ratetypecode = "0000001" Then   'daily rate employee
              mDailyRate = rsEmployee!payrate
            End If
            
            mHrlyRate = Format(mDailyRate / rsParmtr!hrsperday, "#,##0.00")
            
            mMealHrlyRate = Format(rsEmployee!mealallow / rsParmtr!hrsperday, "#,##0.00")
            
            mFEHrlyRate = Format(rsEmployee!fixedearnings / rsParmtr!hrsperday, "#,##0.00")
            
            If rsEmployee!logbased = "Y" Then
                
                NetOpen rsDtrEmp, "select  * from dtremp x1 where x1.employeecode = '" & rsEmployee!employeecode & "' and " & _
                                        "workdate between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' and " & _
                                        "'" & Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "YYYY-MM-DD") & "'"
                
                NetOpen rsDtrEmpTmp, "select  * from dtremp x1 where x1.employeecode = '" & rsEmployee!employeecode & "' and " & _
                                        "workdate between '" & Format(CDate(tdbPayrollPeriod.Columns("wrkdatefrom").Text) - 1, "YYYY-MM-DD") & "' and " & _
                                        "'" & Format(CDate(tdbPayrollPeriod.Columns("wrkdateto").Text) + 1, "YYYY-MM-DD") & "'"
                                        
                With rsDtrEmp
                    
                    If .RecordCount > 0 Then
                    
                        Set rsDtrEmpTmp.DataSource = .Clone
                    
                        Do While Not .EOF
                        
                            mCompHol = True
                        
                            If !Holiday = "Legal" Then
                                
                                If !dayswrk > 0 Then
                                    mDaysWrk = mDaysWrk + 1
                                    mRegAmnt = mRegAmnt + mDailyRate
                                End If
                                
                                rsDtrEmpTmp.Filter = "workdate = '" & CDate(Format(!workdate, "MM/DD/YYYY")) - 1 & "'"
                                
                                If Not rsDtrEmp.EOF Then
                                    If rsDtrEmpTmp!absent = 1 Then
                                        mCompHol = False
                                    End If
                                End If
                                
                                rsDtrEmpTmp.Filter = ""
                                rsDtrEmpTmp.Filter = "workdate = '" & CDate(Format(!workdate, "MM/DD/YYYY")) + 1 & "'"
                            
                                If Not rsDtrEmp.EOF Then
                                    If rsDtrEmpTmp!absent = 1 Then
                                        mCompHol = False
                                    End If
                                End If
                            
                                rsDtrEmpTmp.Filter = ""
                                
                                If mCompHol Then
                                
                                    If !dayswrk > 0 Or !absent > 0 Then
                                        
                                        If !dayoff = "Y" Then
                                        
                                            mRstLegDays = mRstLegDays + 1
                                            mRstLegHrs = mRstLegHrs + !wrkhrs
                                            'mRstLegAmnt = mRstLegAmnt + (!wrkhrs * mHrlyRate * (Format(rsParmtr!restlegholprct / 100, "#,##0.00")))
                                            mRstLegAmnt = mRstLegAmnt + (mDailyRate * (Format(rsParmtr!restlegholprct / 100, "#,##0.00")))
                                            mTtlMealAllow = mTtlMealAllow + (!wrkhrs * mMealHrlyRate * (Format(rsParmtr!restlegholprct / 100, "#,##0.00")))
                                            
                                            mFixedEarnings = mFixedEarnings + (!wrkhrs * mFEHrlyRate * (Format(rsParmtr!restlegholprct / 100, "#,##0.00")))
                                            
                                        Else
                                        
                                            mLegDays = mLegDays + 1
                                            mLegHrs = mLegHrs + !wrkhrs
                                            'mLegAmnt = mLegAmnt + (!wrkhrs * mHrlyRate * (Format(rsParmtr!legholprct / 100, "#,##0.00")))
                                            mLegAmnt = mLegAmnt + (mDailyRate * (Format(rsParmtr!legholprct / 100, "#,##0.00")))
                                            mTtlMealAllow = mTtlMealAllow + (!wrkhrs * mMealHrlyRate * (Format(rsParmtr!legholprct / 100, "#,##0.00")))
                                            
                                            mFixedEarnings = mFixedEarnings + (!wrkhrs * mFEHrlyRate * (Format(rsParmtr!legholprct / 100, "#,##0.00")))
                                            
                                        End If
                                        
                                        mNpLegHrs = mNpLegHrs + !nitewrkhrs
                                        mNpLegAmnt = mNpLegAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!nitelegprct / 100, "#,##0.00")))
    
                                    End If
                                
                                End If
                                
                            ElseIf !Holiday = "Special" Then
                            
                                rsDtrEmpTmp.Filter = "workdate = '" & CDate(Format(!workdate, "MM/DD/YYYY")) - 1 & "'"
                                
                                If Not rsDtrEmp.EOF Then
                                    If rsDtrEmpTmp!absent = 1 Then
                                        mCompHol = False
                                    End If
                                End If
                                
                                rsDtrEmpTmp.Filter = ""
                                rsDtrEmpTmp.Filter = "workdate = '" & CDate(Format(!workdate, "MM/DD/YYYY")) + 1 & "'"
                            
                                If Not rsDtrEmp.EOF Then
                                    If rsDtrEmpTmp!absent = 1 Then
                                        mCompHol = False
                                    End If
                                End If
                                
                                rsDtrEmpTmp.Filter = ""
                                
                                If mCompHol Then
                            
                                    If !dayswrk > 0 Or !absent > 0 Then
                                        
                                        If !dayoff = "Y" Then
                                        
                                            mRstSpcDays = mRstSpcDays + 1
                                            mRstSpcHrs = mRstSpcHrs + !wrkhrs
                                            'mRstSpcAmnt = mRstSpcAmnt + (!wrkhrs * mHrlyRate * (Format(rsParmtr!restspcholprct / 100, "#,##0.00")))
                                            mRstSpcAmnt = mRstSpcAmnt + (mDailyRate * (Format(rsParmtr!restspcholprct / 100, "#,##0.00")))
                                            mTtlMealAllow = mTtlMealAllow + (!wrkhrs * mMealHrlyRate * (Format(rsParmtr!restspcholprct / 100, "#,##0.00")))
                                            
                                            mFixedEarnings = mFixedEarnings + (!wrkhrs * mFEHrlyRate * (Format(rsParmtr!restspcholprct / 100, "#,##0.00")))
                                            
                                        Else
                                        
                                            mSpcDays = mSpcDays + 1
                                            mSpcHrs = mSpcHrs + !wrkhrs
                                            'mSpcAmnt = mSpcAmnt + (!wrkhrs * mHrlyRate * (Format(rsParmtr!spcholprct / 100, "#,##0.00")))
                                            mSpcAmnt = mSpcAmnt + (mDailyRate * (Format(rsParmtr!spcholprct / 100, "#,##0.00")))
                                            
                                            mTtlMealAllow = mTtlMealAllow + (!wrkhrs * mMealHrlyRate * (Format(rsParmtr!spcholprct / 100, "#,##0.00")))
                                            
                                            mFixedEarnings = mFixedEarnings + (!wrkhrs * mFEHrlyRate * (Format(rsParmtr!spcholprct / 100, "#,##0.00")))
                                            
                                        End If
                                        
                                        mNpSpcHrs = mNpSpcHrs + !nitewrkhrs
                                        mNpSpcAmnt = mNpSpcAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!nitespcprct / 100, "#,##0.00")))
                                        If rsEmployee!ratetypecode = "0000004" Then 'for monthly rate
                                            If !dayswrk > 0 Then
                                                mRegAmnt = mRegAmnt - mDailyRate
                                            End If
                                        End If
                                        
                                    End If
                                
                                End If
                                
                            Else
                            
                                If !dayoff = "Y" Then
                                
                                    If !dayswrk > 0 Or !absent > 0 Then
                                    
                                        mRstDays = mRstDays + 1
                                        mRstHrs = mRstHrs + !wrkhrs
                                        mRstAmnt = mRstAmnt + (!wrkhrs * mHrlyRate * (Format(rsParmtr!rstprct / 100, "#,##0.00")))
                                        
                                        If !dayswrk > 0 Then
                                            mTtlMealAllow = mTtlMealAllow + (!wrkhrs * mMealHrlyRate * (Format(rsParmtr!rstprct / 100, "#,##0.00")))
                                            mFixedEarnings = mFixedEarnings + (!wrkhrs * mFEHrlyRate * (Format(rsParmtr!rstprct / 100, "#,##0.00")))
                                        End If
                                        
                                        mNpRstHrs = mNpRstHrs + !nitewrkhrs
                                        mNpRstAmnt = mNpRstAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!niterstprct / 100, "#,##0.00")))
                                        
                                        If rsEmployee!ratetypecode = "0000004" Then 'for monthly rate
                                            If !dayswrk > 0 Then
                                                mRegAmnt = mRegAmnt - mDailyRate
                                            End If
                                        End If
                                        
                                    End If
                                Else
                                
                                    If !dayswrk > 0 Or !absent > 0 Then
                                    
                                        mRegWrkHrs = mRegWrkHrs + !wrkhrs
                                        mAbsDays = mAbsDays + !absent
                                        mLatehrs = mLatehrs + !latehrs
                                        mUtHrs = mUtHrs + !uthrs
                                        mLateAmnt = mLateAmnt + (!latehrs * mHrlyRate)
                                        mUtAmnt = mUtHrs + (!uthrs * mHrlyRate)
                                        
                                        If rsEmployee!ratetypecode = "0000001" Then 'for daily rate employees
                                            mDaysWrk = mDaysWrk + !dayswrk
                                            If !dayswrk > 0 Then
                                                mRegAmnt = mRegAmnt + rsEmployee!payrate
                                            End If
                                        Else
                                            mDaysWrk = mDaysWrk + 1
                                            'mRegAmnt = mRegAmnt + rsEmployee!payrate
                                            mAbsAmnt = mAbsAmnt + (!absent * mDailyRate)
                                        End If
                                        
                                        If !dayswrk > 0 Then
                                            mFixedEarnings = mFixedEarnings + rsEmployee!fixedearnings
                                            mTtlMealAllow = mTtlMealAllow + rsEmployee!mealallow
                                        End If
                                        
                                        mNpRegHrs = mNpRegHrs + !nitewrkhrs
                                        mNpRegAmnt = mNpRegAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!niteregprct / 100, "#,##0.00")))
                                    
                                    End If
                                    
                                End If
                                
                            End If
                            
                           .MoveNext
                           
                        Loop
                        
                    End If
                    
                End With
                
            Else
                
                mRegAmnt = 0
            
            End If
            
            NetOpen rsOT, "select * from overtimelne where fnlz <> 'Y' and employeecode = '" & rsEmployee!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "' and status = 'Approved'"
            
            With rsOT
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                    
                        If !Holiday = "Legal" Then
                            mOtLegHrs = mOtLegHrs + !otwrkhrs
                            mOtNpLegAmnt = mOtNpLegAmnt + !nitewrkhrs
                            mOtLegAmnt = mOtLegAmnt + (!otwrkhrs * mHrlyRate * (Format(rsParmtr!otlegprct / 100, "#,##0.00")))
                            mOtNpLegAmnt = mOtNpLegAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!otnitelegprct / 100, "#,##0.00")))
                            
                            
                            If !dayoff = "Y" Then
                                mRstLegHrs = mRstLegHrs + !regwrkhrs
                                mRstLegAmnt = mRstLegAmnt + (!regwrkhrs * mHrlyRate * (Format(rsParmtr!restlegholprct / 100, "#,##0.00")))
                            Else
                                mLegHrs = mLegHrs + !regwrkhrs
                                mLegAmnt = mLegAmnt + (!regwrkhrs * mHrlyRate * (Format(rsParmtr!legholprct / 100, "#,##0.00")))
                            End If
                            
                        ElseIf !Holiday = "Special" Then
                            
                            mOtSpcHrs = mOtSpcHrs + !otwrkhrs
                            mOtNpSpcHrs = mOtNpSpcHrs + !nitewrkhrs
                            mOtSpcAmnt = mOtSpcAmnt + (!otwrkhrs * mHrlyRate * (Format(rsParmtr!otspcprct / 100, "#,##0.00")))
                            mOtNpSpcAmnt = mOtNpSpcAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!otnitespcprct / 100, "#,##0.00")))
                            
                            If !dayoff = "Y" Then
                                mRstSpcHrs = mRstSpcHrs + !regwrkhrs
                                mRstSpcAmnt = mRstSpcAmnt + (!regwrkhrs * mHrlyRate * (Format(rsParmtr!restspcholprct / 100, "#,##0.00")))
                            Else
                                mSpcHrs = mSpcHrs + !wrkhrs
                                mSpcAmnt = mSpcAmnt + (!regwrkhrs * mHrlyRate * (Format(rsParmtr!spcholprct / 100, "#,##0.00")))
                            End If
                            
                        Else
                            If !dayoff = "Y" Then
                                mOtRstHrs = mOtRstHrs + !otwrkhrs
                                mOtNpRstHrs = mOtNpRstHrs + !nitewrkhrs
                                mOtRstAmnt = mOtRstAmnt + (!otwrkhrs * mHrlyRate * (Format(rsParmtr!otrstprct / 100, "#,##0.00")))
                                mOtNpRstAmnt = mOtNpRstAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!otniterstprct / 100, "#,##0.00")))
                                
                                mRstHrs = mRstHrs + !regwrkhrs
                                mRstAmnt = mRstAmnt + (!regwrkhrs * mHrlyRate * (Format(rsParmtr!rstprct / 100, "#,##0.00")))
                                
                            Else
                                mOtRegHrs = mOtRegHrs + !otwrkhrs
                                mOtNpRegHrs = mOtNpRegHrs + !nitewrkhrs
                                mOtRegAmnt = mOtRegAmnt + (!otwrkhrs * mHrlyRate * (Format(rsParmtr!otregprct / 100, "#,##0.00")))
                                mOtNpRegAmnt = mOtNpRegAmnt + (!nitewrkhrs * mHrlyRate * (Format(rsParmtr!otniteregprct / 100, "#,##0.00")))
                                
                                mRegWrkHrs = mRegWrkHrs + !regwrkhrs
                                mRegAmnt = mRegAmnt + (!regwrkhrs * mHrlyRate)
                                
                            End If
                        End If
                        .MoveNext
                    Loop
                End If
            End With
            
            
            NetOpen rsLeave, "select x1.* from lvlne x1 " & _
                                 "left outer join lvhdr x2 on x1.lvnum = x2.lvnum " & _
                                 "where x2.cancel <> 'Y' and x2.employeecode = '" & rsEmployee!employeecode & "' and " & _
                                 "X1.lvdate between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' and " & _
                                 "'" & Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "YYYY-MM-DD") & "'"
                                 
            With rsLeave
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If !withpay = "Y" Then
                            If !firstshift = 1 Then
                                mLvWPDays = mLvWPDays + 0.5
                            End If
                            If !secondshift = 1 Then
                                mLvWPDays = mLvWPDays + 0.5
                            End If
                        Else
                            mLvWoPDays = mLvWoPDays + 1
                        End If
                        .MoveNext
                    Loop
                End If
                mLvWPAmnt = mLvWPDays * mDailyRate
            End With
            
            NetOpen rsEarnings, "select sum(amount) amount from earnings where percode = '" & tdbPayrollPeriod.BoundText & "' and employeecode = '" & rsEmployee!employeecode & "'"
            
            With rsEarnings
                If .RecordCount > 0 Then
                    If !amount > 0 Then
                        mEarnings = !amount
                    End If
                End If
            End With
            
            
            mRegAmnt = Format(mRegAmnt, "#,##0.00")
            mPORegAmnt = Format(mPORegAmnt, "#,##0.00")
            mPOSpcAmnt = Format(mPOSpcAmnt, "#,##0.00")
            mPORstAmnt = Format(mPORstAmnt, "#,##0.00")
            mPOLegAmnt = Format(mPOLegAmnt, "#,##0.00")
            mPONpRegAmnt = Format(mPONpRegAmnt, "#,##0.00")
            mPONpSpcAmnt = Format(mPONpSpcAmnt, "#,##0.00")
            mPONpRstAmnt = Format(mPONpRstAmnt, "#,##0.00")
            mPONpLegAmnt = Format(mPONpLegAmnt, "#,##0.00")
            mAbsAmnt = Format(mAbsAmnt, "#,##0.00")
            mLateAmnt = Format(mLateAmnt, "#,##0.00")
            mUtAmnt = Format(mUtAmnt, "#,##0.00")
            mLegAmnt = Format(mLegAmnt, "#,##0.00")
            mSpcAmnt = Format(mSpcAmnt, "#,##0.00")
            mRstLegAmnt = Format(mRstLegAmnt, "#,##0.00")
            mRstSpcAmnt = Format(mRstSpcAmnt, "#,##0.00")
            mRstAmnt = Format(mRstAmnt, "#,##0.00")
            mOtRegAmnt = Format(mOtRegAmnt, "#,##0.00")
            mOtSpcAmnt = Format(mOtSpcAmnt, "#,##0.00")
            mOtRstAmnt = Format(mOtRstAmnt, "#,##0.00")
            mOtLegAmnt = Format(mOtLegAmnt, "#,##0.00")
            mOtNpRegAmnt = Format(mOtNpRegAmnt, "#,##0.00")
            mOtNpSpcAmnt = Format(mOtNpSpcAmnt, "#,##0.00")
            mOtNpRstAmnt = Format(mOtNpRstAmnt, "#,##0.00")
            mOtNpLegAmnt = Format(mOtNpLegAmnt, "#,##0.00")
            mNpRegAmnt = Format(mNpRegAmnt, "#,##0.00")
            mNpSpcAmnt = Format(mNpSpcAmnt, "#,##0.00")
            mNpRstAmnt = Format(mNpRstAmnt, "#,##0.00")
            mNpLegAmnt = Format(mNpLegAmnt, "#,##0.00")
            mEarnings = Format(mEarnings, "#,##0.00")
            mLvWPAmnt = Format(mLvWPAmnt, "#,##0.00")
            
            mBasic = (mRegAmnt + mPORegAmnt + mPOSpcAmnt + mPORstAmnt + mPOLegAmnt + mPONpRegAmnt + mPONpRstAmnt + _
                                 mPONpSpcAmnt + mPONpLegAmnt + mFixedEarnings) - (mAbsAmnt + mLateAmnt + mUtAmnt + mTtlMealAllow)
            
            mGross = mBasic + mLegAmnt + mSpcAmnt + mRstLegAmnt + mRstSpcAmnt + mRstAmnt + _
                    mOtRegAmnt + mOtRstAmnt + mOtSpcAmnt + mOtLegAmnt + _
                    mNpRegAmnt + mNpRstAmnt + mNpSpcAmnt + mNpLegAmnt + _
                    mOtNpRegAmnt + mOtNpRstAmnt + mOtNpSpcAmnt + mOtNpLegAmnt + mEarnings + mLvWPAmnt
            
            NetOpen rsDeductions, "select sum(amount) amount from deductions where percode = '" & tdbPayrollPeriod.BoundText & "' and employeecode = '" & rsEmployee!employeecode & "'"
            With rsDeductions
                If .RecordCount > 0 Then
                    If !amount > 0 Then
                        mDeductions = !amount
                    End If
                End If
            End With
            
            NetOpen rsLoanDed, "select sum(amtded) amount from loanded where fnlz <> 'Y' and " & _
                        "percode = '" & tdbPayrollPeriod.BoundText & "' and employeecode = '" & rsEmployee!employeecode & "'"
            With rsLoanDed
                If .RecordCount > 0 Then
                    If !amount > 0 Then
                        mLoanded = !amount
                    End If
                End If
            End With
            
            If rsEmployee!sssauto = 0 Then
                mSssAmnt = rsEmployee!sssamt
                mSSSEr = rsEmployee!ssser
                mSSSEc = rsEmployee!sssec
            Else
                If rsEmployee!ratetypecode = "0000004" Then
                    CompSss rsEmployee!payrate, mSssAmnt, mSSSEr, mSSSEc
                ElseIf rsEmployee!ratetypecode = "0000001" Then
                    CompSss rsEmployee!payrate * (314 / 12), mSssAmnt, mSSSEr, mSSSEc
                End If
            End If
            
            If rsEmployee!philhauto = 0 Then
                mPhilAmnt = rsEmployee!philhamt
                mPhilEr = rsEmployee!philher
            Else
                If rsEmployee!ratetypecode = "0000004" Then
                    CompPhi rsEmployee!payrate, mPhilAmnt, mPhilEr
                ElseIf rsEmployee!ratetypecode = "0000001" Then
                    CompPhi rsEmployee!payrate * (314 / 12), mPhilAmnt, mPhilEr
                End If
            End If
            
            If rsEmployee!hdmfauto = 0 Then
                mHdmfAmnt = rsEmployee!hdmfamt
                mHdmfEr = rsEmployee!hdmfer
            End If
            
            If rsEmployee!taxauto = 0 Then
                mTaxAmnt = rsEmployee!taxamt
            Else
                CompTax mGross - (mSssAmnt + mPhilAmnt + mHdmfAmnt), rsEmployee!wtcode, mTaxAmnt
            End If
            
            mSssAmnt = Format(mSssAmnt, "#,##0.00")
            mPhilAmnt = Format(mPhilAmnt, "#,##0.00")
            mHdmfAmnt = Format(mHdmfAmnt, "#,##0.00")
            mTaxAmnt = Format(mTaxAmnt, "#,##0.00")
            mDeductions = Format(mDeductions, "#,##0.00")
            mLoanded = Format(mLoanded, "#,##0.00")
            
            mNet = mGross - (mSssAmnt + mPhilAmnt + mHdmfAmnt + mTaxAmnt + mDeductions + mLoanded)
            
            ConMain.Execute "delete from payroll where employeecode = '" & rsEmployee!employeecode & "' and percode = '" & tdbPayrollPeriod.BoundText & "'"
            
            ConMain.Execute "insert into payroll (percode,branchcode,divisioncode,costcentercode,sectioncode,employeecode,ratetypecode,regamnt,dayswrk, " & _
                        "absdays,absamnt, regwrkhrs,latehrs,lateamnt,uthrs,utamnt,legdays,leghrs,legamnt,spcdays,spchrs,spcamnt,rstdays,rsthrs,rstamnt, " & _
                        "rstlegdays,rstleghrs,rstlegamnt,rstspcdays,rstspchrs,rstspcamnt, " & _
                        "otreghrs,otregamnt,otleghrs,otlegamnt,otspchrs,otspcamnt,otrsthrs,otrstamnt, " & _
                        "npreghrs,npregamnt,npspchrs,npspcamnt,npleghrs,nplegamnt,nprsthrs,nprstamnt," & _
                        "otnpreghrs,otnpregamnt,otnpspchrs,otnpspcamnt,otnpleghrs,otnplegamnt,otnprsthrs,otnprstamnt," & _
                        "poreghrs,poregamnt,pospchrs,pospcamnt,poleghrs,polegamnt,porsthrs,porstamnt," & _
                        "ponpreghrs,ponpregamnt,ponpspchrs,ponpspcamnt,ponpleghrs,ponplegamnt,ponprsthrs,ponprstamnt," & _
                        "earnings,gross,sssamnt,ssser,ec,philamnt,philer,hdmfamnt,payrate,bankacctno,saltobank,regular, " & _
                        "hdmfer,taxamnt, deductions,net,basic,lvwpdays,lvwpamnt,lvwopdays,lvwopamnt,loanded,payyear,paymonth,mealallow,fixedearnings) values " & _
                        "('" & tdbPayrollPeriod.BoundText & "','" & rsEmployee!branchcode & "','" & rsEmployee!divisioncode & "','" & rsEmployee!costcentercode & "','" & rsEmployee!sectioncode & "','" & rsEmployee!employeecode & "','" & rsEmployee!ratetypecode & "'," & mRegAmnt & "," & mDaysWrk & ", " & _
                        mAbsDays & "," & mAbsAmnt & "," & mRegWrkHrs & "," & mLatehrs & "," & mLateAmnt & "," & mUtHrs & "," & mUtAmnt & "," & mLegDays & "," & mLegHrs & "," & mLegAmnt & "," & mSpcDays & "," & mSpcHrs & "," & mSpcAmnt & "," & mRstDays & "," & mRstHrs & "," & mRstAmnt & ", " & _
                        mRstLegDays & "," & mRstLegHrs & "," & mRstLegAmnt & "," & mRstSpcDays & "," & mRstSpcHrs & "," & mRstSpcAmnt & ", " & _
                        mOtRegHrs & ", " & mOtRegAmnt & ", " & mOtLegHrs & "," & mOtLegAmnt & "," & mOtSpcHrs & "," & mOtSpcAmnt & "," & mOtRstHrs & "," & mOtRstAmnt & ", " & _
                        mNpRegHrs & "," & mNpRegAmnt & "," & mNpSpcHrs & "," & mNpSpcAmnt & "," & mNpLegHrs & "," & mNpLegAmnt & "," & mNpRstHrs & "," & mNpRstAmnt & ", " & _
                        mOtNpRegHrs & "," & mOtNpRegAmnt & "," & mOtNpSpcHrs & "," & mOtNpSpcAmnt & "," & mOtNpLegHrs & "," & mOtNpLegAmnt & "," & mOtNpRstHrs & "," & mOtNpRstAmnt & ", " & _
                        mPORegHrs & "," & mPORegAmnt & "," & mPOSpcHrs & "," & mPOSpcAmnt & "," & mPOLegHrs & "," & mPOLegAmnt & "," & mPORstHrs & "," & mPORstAmnt & ", " & _
                        mPONpRegHrs & "," & mPONpRegAmnt & "," & mPONpSpcHrs & "," & mPONpSpcAmnt & "," & mPONpLegHrs & "," & mPONpLegAmnt & "," & mPONpRstHrs & "," & mPONpRstAmnt & ", " & _
                        mEarnings & "," & mGross & "," & mSssAmnt & "," & mSSSEr & "," & mSSSEc & "," & mPhilAmnt & "," & mPhilEr & "," & mHdmfAmnt & "," & rsEmployee!payrate & ",'" & rsEmployee!bankacctno & "','" & rsEmployee!saltobank & "','" & rsEmployee!regular & "', " & _
                        mHdmfEr & "," & mTaxAmnt & "," & mDeductions & "," & mNet & "," & mBasic & "," & mLvWPDays & "," & mLvWPAmnt & "," & mLvWoPDays & "," & mLvWoPAmnt & "," & mLoanded & ",'" & tdbPayrollPeriod.Columns("payyear").Text & "','" & tdbPayrollPeriod.Columns("paymonth").Text & "'," & mTtlMealAllow & "," & mFixedEarnings & ")"
            
            rsEmployee.MoveNext
            DoEvents
        Loop
                
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

Private Sub CompSss(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double, ByRef mEc As Double)

    Dim rsTmp   As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    
    NetOpen rsTmp, "select * from sss order by fromamount"
    
    If rsTmp.RecordCount > 0 Then
      With rsTmp
        .MoveFirst
        Do While Not .EOF
          If (par1 >= .Fields("fromamount")) And (par1 <= .Fields("toamount")) Then
            mAmount = .Fields("ee")
            mEr = .Fields("er")
            mEc = .Fields("ec")
            Exit Do
          End If
          If .AbsolutePosition = .RecordCount Then
            mAmount = .Fields("ee")
            mEr = .Fields("er")
            mEc = .Fields("ec")
            Exit Do
          End If
          .MoveNext
        Loop
      End With
    End If
    
End Sub

Private Sub CompPhi(ByVal par1 As Double, ByRef mAmount As Double, ByRef mEr As Double)
        
    Dim rsTmp   As ADODB.Recordset
    Set rsTmp = New ADODB.Recordset
    NetOpen rsTmp, "select * from ph order by fromamount"
    If rsTmp.RecordCount > 0 Then
        With rsTmp
          .MoveFirst
          Do While Not .EOF
            If (par1 >= .Fields("fromamount")) And (par1 <= .Fields("toamount")) Then
              mAmount = .Fields("ee")
              mEr = .Fields("er")
              Exit Do
            End If
            If .AbsolutePosition = .RecordCount Then
              mAmount = .Fields("ee")
              mEr = .Fields("er")
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






