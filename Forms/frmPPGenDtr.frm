VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPGenDtr 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate DTR"
   ClientHeight    =   3060
   ClientLeft      =   6135
   ClientTop       =   4635
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPPGenDtr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6165
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
      Height          =   3585
      Left            =   -15
      TabIndex        =   6
      Top             =   -90
      Width           =   6240
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1980
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   285
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
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
         _PropDict       =   $"frmPPGenDtr.frx":6852
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
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   330
         Left            =   135
         TabIndex        =   5
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
         TabIndex        =   12
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
         TabIndex        =   4
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
         _PropDict       =   $"frmPPGenDtr.frx":68FC
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
         _PropDict       =   $"frmPPGenDtr.frx":69A6
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
         TabIndex        =   3
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
         _PropDict       =   $"frmPPGenDtr.frx":6A50
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
         _PropDict       =   $"frmPPGenDtr.frx":6AFA
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
         TabIndex        =   13
         Top             =   2880
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0030A0B8&
         X1              =   120
         X2              =   6000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0030A0B8&
         X1              =   120
         X2              =   6000
         Y1              =   750
         Y2              =   750
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
         TabIndex        =   11
         Top             =   930
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
         TabIndex        =   10
         Top             =   1305
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
         TabIndex        =   9
         Top             =   1665
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
         ForeColor       =   &H0030A0B8&
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   345
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
         ForeColor       =   &H0030A0B8&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   2235
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frmPPGenDtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTitoTmp       As ADODB.Recordset
Dim rsDtrTmp        As ADODB.Recordset
Dim rsEmployee      As ADODB.Recordset
Dim rsOTTmp         As ADODB.Recordset
Dim rsDTSTmp        As ADODB.Recordset

Private Sub Form_Load()
    
    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    SendMessage pb2.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb2.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    bind_tdb ConMain, tdbPayrollPeriod, "select percode, description, wrkdatefrom, wrkdateto, payfreqcode from payrollperiod", "description", "percode"
    bind_tdb ConMain, tdbBranch, "select branchcode, branch from branch order by branch", "branch", "branchcode"
      
End Sub

Private Sub Form_Resize()

    On Error Resume Next
        
End Sub

Private Sub cmdGenerate_Click()

    Dim mTin            As Date
    Dim mTout           As Date
    Dim mLasTout        As String
    Dim mAdvDate        As String
    
    Dim mT1in           As String
    Dim mT1out          As String
    Dim mT2in           As String
    Dim mT2out          As String
    Dim mST1in          As String
    Dim mST1out         As String
    Dim mST2in          As String
    Dim mST2out         As String
    Dim mWrkdate        As String

    Dim rsOT            As ADODB.Recordset
    Dim rsDtrEmp        As ADODB.Recordset
    Dim rsHoliday       As ADODB.Recordset
    Dim rsParmtr        As ADODB.Recordset
    Dim rsDts           As ADODB.Recordset
    Dim rsOTChk         As ADODB.Recordset
    Dim rsOTCmpr        As ADODB.Recordset
    Dim rsCostCenter    As ADODB.Recordset
    
    Dim mOT_Start       As String
    Dim mOT_End         As String
    Dim mNiteStart      As String
    Dim mNiteEnd        As String
    Dim mTimeIN         As String
    Dim mTimeOUT        As String
    Dim mDTS_Start      As String
    Dim mDTS_End        As String
    
    Dim mDTSActIN       As Date
    Dim mDTSActOUT      As Date
    Dim mOTActIN        As Date
    Dim mOTActOUT       As Date
    Dim mNiteIN         As Date
    Dim mNiteOUT        As Date
    Dim mPrevOUT        As Date
    Dim mTimeFrom       As Date
    Dim mTimeTo         As Date
    
    Dim mLateAllow      As Double
    Dim mOTAllow        As Double
    Dim mWrkHrs         As Double
    Dim mNiteWrkHrs     As Double
    Dim mUTMin          As Double
    Dim mLackHrs        As Double
    Dim mTtlReghrs      As Double
    
'*******************************************************
    
    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        MsgBox "Please select a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Do you want to proceed in generating DTRs?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If Trim(tdbBranch.Text) <> "" And Not IsNull(tdbBranch.SelectedItem) And tdbBranch.ApproxCount > 0 Then
        If Trim(tdbDivision.Text) <> "" And Not IsNull(tdbDivision.SelectedItem) And tdbDivision.ApproxCount > 0 Then
            If Trim(tdbCostCenter.Text) <> "" And Not IsNull(tdbCostCenter.SelectedItem) And tdbCostCenter.ApproxCount > 0 Then
                If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
                    NetOpen rsEmployee, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' and  " & _
                                            "costcentercode = '" & tdbCostCenter.BoundText & "' and employeecode = '" & tdbEmployee.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
                Else
                    NetOpen rsEmployee, "select employeecode,branchcode, divisioncode, costcentercode,sectioncode from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' and " & _
                                            "costcentercode = '" & tdbCostCenter.BoundText & "' and " & _
                                            "payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
                End If
            Else
                NetOpen rsEmployee, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and divisioncode = '" & tdbDivision.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
            End If
        Else
            NetOpen rsEmployee, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee where branchcode = '" & tdbBranch.BoundText & "' " & _
                                            "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
        End If
    Else
        NetOpen rsEmployee, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee where payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' "
    End If
    
    If rsEmployee.RecordCount > 0 Then
        
        fra1.Enabled = False
        Me.MousePointer = vbHourglass
        cmdGenerate.Enabled = False
        
        pb1.Max = rsEmployee.RecordCount
        pb1.Value = 0
        
        rsEmployee.MoveFirst
        
        NetOpen rsParmtr, "select * from parmtr"
        If rsParmtr.RecordCount > 0 Then
            mLateAllow = Format(rsParmtr!lateallowance / 60, "#,##0.00")
            mUTMin = Format(rsParmtr!utmin / 60, "#,##0.00")
        End If
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        
        Do While Not rsEmployee.EOF
            
            pb1.Value = pb1.Value + 1
                                                    
            ConMain.Execute "delete from dtremp where employeecode = '" & rsEmployee!employeecode & "' and " & _
                                    "workdate between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "'  and  " & _
                                    "'" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' and updatable = 'Y'"
                                                    
            Create_TmpTito
            
            Load_Dtr
            
            mLasTout = ""
            
            If rsDtrTmp.RecordCount > 0 Then
                'Assigning TITO to Employee DTR
                
                With rsDtrTmp
                    pb2.Max = .RecordCount
                    pb2.Value = 0
                    .MoveFirst
                    Do While Not .EOF
                        pb2.Value = pb2.Value + 1
                        
                        !dayswrk = 0
                        !wrkhrs = 0
                        !nitewrkhrs = 0
                        !absent = 0
                        !latehrs = 0
                        !uthrs = 0
                        
                        If rsTitoTmp.RecordCount > 0 Then
                            If !updatable <> 0 Then
                                'For shifts that only have two time slots
                                If Trim(!st1in) <> "" And Trim(!st1out) <> "" And Trim(!st2in) = "" And Trim(!st2out) = "" Then
                                    mTin = Format(CDate(!wrkdate) & " " & !st1in, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    If CDate(!st1in) > CDate(!st1out) Then
                                        mTout = Format(CDate(!wrkdate) + 1 & " " & !st1out, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    Else
                                        mTout = Format(CDate(!wrkdate) & " " & !st1out, "MM/DD/YYYY hh:nn:ss AM/PM")
                                    End If
                                    mST1in = mTin
                                    mST1out = mTout
                                    If mLasTout = "" Then
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "'"
                                    Else
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    End If
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t1in) = "" Then !t1in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t1out) = "" Then !t1out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                End If
                                
                                'For shifts that have four time slots
                                If Trim(!st1in) <> "" And Trim(!st1out) <> "" And Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                    mTin = Format(CDate(!wrkdate) & " " & !st1in, "MM/DD/YYYY HH:NN:SS AM/PM")
                                    If CDate(!st1in) > CDate(!st1out) Then
                                        mAdvDate = CDate(!wrkdate) + 1
                                    Else
                                        mAdvDate = CDate(!wrkdate)
                                    End If
                                    mTout = Format(CDate(mAdvDate) & " " & !st1out, "MM/DD/YYYY HH:NN:S AM/PM")
                                    mST1in = mTin
                                    mST1out = mTout
                                    If mLasTout = "" Then
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "'"
                                    Else
                                        rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    End If
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t1in) = "" Then !t1in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t1out) = "" Then !t1out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                    If CDate(!st1out) > CDate(!st2in) Then
                                        mAdvDate = CDate(mAdvDate) + 1
                                    Else
                                        mAdvDate = CDate(mAdvDate)
                                    End If
                                    mTin = Format(CDate(mAdvDate) & " " & !st2in, "MM/DD/YYYY HH:NN:SS AM/PM")
                                    
                                    If CDate(!st2in) > CDate(!st2out) Then
                                        mAdvDate = CDate(mAdvDate) + 1
                                    Else
                                        mAdvDate = CDate(mAdvDate)
                                    End If
                                    
                                    mTout = Format(CDate(mAdvDate) & " " & !st2out, "MM/DD/YYYY HH:NN:S AM/PM")
                                    mST2in = mTin
                                    mST2out = mTout
                                    rsTitoTmp.Filter = "tout > '" & mTin & "' and tin < '" & mTout & "' and tin > '" & mLasTout & "'"
                                    If rsTitoTmp.RecordCount > 0 Then
                                        rsTitoTmp.MoveFirst
                                        Do While Not rsTitoTmp.EOF
                                            If rsTitoTmp.AbsolutePosition = 1 Then
                                                If Trim(!t2in) = "" Then !t2in = Format(rsTitoTmp!tin, "hh:nn:ss")
                                            End If
                                            If rsTitoTmp.AbsolutePosition = rsTitoTmp.RecordCount Then
                                                If Trim(!t2out) = "" Then !t2out = Format(rsTitoTmp!tout, "hh:nn:ss")
                                            End If
                                            rsTitoTmp.MoveNext
                                        Loop
                                    End If
                                    
                                    mLasTout = mTout
                                    rsTitoTmp.Filter = ""
                                    
                                End If
                            End If
                        End If
                    
                        'Do the computations
                        'Clear all Time and Date variables
                        
                        mT1in = ""
                        mT1out = ""
                        mT2in = ""
                        mT2out = ""
                        mST1in = ""
                        mST1out = ""
                        mST2in = ""
                        mST2out = ""
                        mWrkHrs = 0
                        mNiteWrkHrs = 0
                        
                        'Check if employee has a schedule for today
                        If Trim(!st1in) <> "" And Trim(!st1out) <> "" Then
                            'Set Workdate to be used for shifting schedule
                            mWrkdate = !wrkdate
                            'assgin shifting schedule variables with vaues
                            mST1in = mWrkdate & " " & !st1in
                            If CDate(mST1in) > CDate(mWrkdate & " " & !st1out) Then
                                mWrkdate = Format(CDate(!wrkdate) + 1, "MM/DD/YYYY")
                            End If
                            
                            mST1out = mWrkdate & " " & !st1out
                            If Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                If CDate(mST1out) > CDate(mWrkdate & " " & !st2in) Then
                                    mWrkdate = CDate(mWrkdate) + 1
                                End If
                                mST2in = mWrkdate & "  " & !st2in
                                If CDate(mST2in) > CDate(mWrkdate & " " & !st2out) Then
                                    mWrkdate = CDate(mWrkdate + 1)
                                End If
                                mST2out = mWrkdate & " " & !st2out
                            End If
                            'Set Workdate back to original date to be used for actual time logs
                            mWrkdate = !wrkdate
                            'set values for 1st Tito
                            If Trim(!t1in) <> "" And Trim(!t1out) <> "" Then
                                If CDate(mWrkdate & " " & !t1in) > CDate(mST1out) Then
                                    mT1in = CDate(mWrkdate) - 1 & " " & !t1in
                                Else
                                    mT1in = mWrkdate & " " & !t1in
                                End If
                                If CDate(mT1in) > CDate(mWrkdate & " " & !t1out) Then
                                    mWrkdate = CDate(mWrkdate) + 1
                                End If
                                mT1out = mWrkdate & " " & !t1out
                                If CDate(mT1out) < CDate(mST1in) And CDate(mT1out) < CDate(mST1out) Then
                                        mWrkdate = CDate(mWrkdate) + 1
                                        mT1out = mWrkdate & " " & !t1out
                                End If
                            End If
                            
                            'set values for 2nd tito
                            If Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                                If Trim(!t2in) <> "" And Trim(!t2out) <> "" Then
                                    If Trim(!t1in) <> "" Then
                                        If CDate(mT1out) > CDate(mWrkdate & " " & !t2in) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2in = mWrkdate & " " & !t2in
                                        If CDate(mT2in) > CDate(mWrkdate & " " & !t2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2out = mWrkdate & " " & !t2out
                                    Else
                                        If CDate(mWrkdate & " " & !t2in) > CDate(mST2out) Then
                                            mT2in = CDate(mWrkdate) - 1 & " " & !t2in
                                        Else
                                            mT2in = mWrkdate & " " & !t2in
                                        End If
                                        If CDate(mT2in) > CDate(mWrkdate & " " & !t2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                        End If
                                        mT2out = mWrkdate & " " & !t2out
                                    End If
                                    If CDate(mT2out) < CDate(mST2in) And CDate(mT2out) < CDate(mST2out) Then
                                            mWrkdate = CDate(mWrkdate) + 1
                                            mT2out = mWrkdate & " " & !t2out
                                    End If
                                End If
                            End If
                            
                            'Compute number of hours work, lates and undertimes.
                            If mST2in <> "" Then 'For schedules with four(4) time slots
                                If !travel = 0 And !leave = 0 Then
                                    If mT1in <> "" And mT2in <> "" Then 'if all four(4) time slots were consumed.
                                        !dayswrk = 1
                                        'compute for late
                                        If CDate(mT1in) > CDate(mST1in) Then
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) - mLateAllow
                                                End If
                                            End If
                                        End If
                                        'check if late on 2nd time in
                                        If !brkhrsperday > 0 Then
                                            If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                                If (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) > mLateAllow Then
                                                    If .Fields("latehrs") > 0 Then
                                                        .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) - mLateAllow
                                                    Else
                                                        .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday - mLateAllow
                                                    End If
                                                End If
                                            End If
                                        End If
                                        'compute for undertime
                                        If CDate(mT2out) < CDate(mST2out) Then
                                            If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                            End If
                                        End If
                                        mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                        mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                        mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    ElseIf mT1in <> "" And mT2in = "" Then 'if only the first two (2) time slots were consumed.
                                    
                                        If CDate(mT1in) <= CDate(mST1in) Then
                                            mTin = mST1in
                                        Else
                                            mTin = mT1in
                                            'compute for late
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) - mLateAllow
                                                End If
                                            End If
                                        End If
                                        
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            mTout = mT1out
                                            mWrkHrs = DiffHrs(mTin, mTout)
                                            'compute for 1st undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                            If Trim(.Fields("holiday")) = "" Then
                                                .Fields("absent") = 0.5
                                            End If
                                            .Fields("dayswrk") = 0.5
                                        ElseIf CDate(mT1out) >= CDate(mST1out) And CDate(mT1out) < CDate(mST2in) Then
                                            mTout = mST1out
                                            mWrkHrs = DiffHrs(mTin, mTout)
                                            If Trim(.Fields("holiday")) = "" Then
                                                .Fields("absent") = 0.5
                                            End If
                                            .Fields("dayswrk") = 0.5
                                        ElseIf CDate(mT1out) >= CDate(mST2in) And CDate(mT1out) < CDate(mST2out) Then
                                            mTout = mT1out
                                            'compute for 2nd undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST2out))
                                            End If
                                            mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                            .Fields("dayswrk") = 1
                                        Else
                                            mTout = mST2out
                                            mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                            .Fields("dayswrk") = 1
                                        End If
                                        
                                        mNiteWrkHrs = DiffNiteHrs(mTin, mTout, mST1in, mST2out, !nitepremstart, !nitepremend)
                                    ElseIf mT1in = "" And mT2in <> "" Then ' if only the last two(2) time slots were consumed.
                                        If CDate(mT2in) <= CDate(mST2in) Then
                                            mTin = mST2in
                                        Else
                                            mTin = mT2in
                                            If DiffHrs(CDate(mST1in), CDate(mT2in)) > 0 Then
                                                If DiffHrs(CDate(mST2in), CDate(mT2in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in)) - mLateAllow
                                                End If
                                            End If
                                        End If
                                        If CDate(mT2out) < CDate(mST2out) Then
                                            mTout = mT2out
                                             'compute for undertime
                                            If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                            End If
                                        Else
                                            mTout = mST2out
                                        End If
                                        If Trim(.Fields("holiday")) = "" Then
                                            .Fields("absent") = 0.5
                                        End If
                                        .Fields("dayswrk") = 0.5
                                        mNiteWrkHrs = DiffNiteHrs(mTin, mTout, mST2in, mST2out, !nitepremstart, !nitepremend)
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                    Else
                                        If Trim(.Fields("holiday")) = "" Then
                                            If .Fields("required") = "Y" Then
                                                .Fields("absent") = 1
                                            Else
                                                .Fields("absent") = 0
                                            End If
                                            .Fields("dayswrk") = 0
                                        End If
                                    End If
                                    
                                Else 'On travel or On leave
                                    'If on travel or on leave during the second shift, compute only the first shift.
                                    If !firsttravel = 0 And !firstleave = 0 Then
                                        If mT1in <> "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT1in) > CDate(mST1in) Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                    If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) - mLateAllow
                                                    End If
                                                End If
                                            End If
                                            If CDate(mT1out) < CDate(mST1out) Then
                                                If CDate(mT2in) < CDate(mST1out) Then
                                                    'compute for late
                                                    If !brkhrsperday > 0 Then
                                                        If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                                            If (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) > mLateAllow Then
                                                                If .Fields("latehrs") > 0 Then
                                                                    .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday) - mLateAllow
                                                                Else
                                                                    .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday - mLateAllow
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    'compute for undertime
                                                    If CDate(mT1out) < CDate(mST1out) Then
                                                        If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                            .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !secondtravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mST2in, mST1in, mT1out, !nitepremstart, !nitepremend)
                                                mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mST2in, mST2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in <> "" And mT2in = "" Then
                                            'compute for late
                                            If CDate(mT1in) > CDate(mST1in) Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                    If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) - mLateAllow
                                                    End If
                                                End If
                                            End If
                                            'Compute for undertime
                                            If CDate(mT1out) < CDate(mST1out) Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !secondtravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in = "" And mT2in <> "" Then
                                            .Fields("absent") = 0.5
                                            .Fields("dayswrk") = 0.5
                                            If !secondtravel = 1 Then
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                            End If
                                        End If
                                        
                                    'If on travel or on leave during the first shift, compute only the second shift
                                    ElseIf !secondtravel = 0 And !secondleave = 0 Then
                                        If mT1in <> "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT1out) >= CDate(mST2in) Then
                                                If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                    If DiffHrs(CDate(mT1out), CDate(mT2in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - mLateAllow
                                                    End If
                                                End If
                                            Else
                                                If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                    If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                        If DiffHrs(CDate(mT2in), CDate(mST1out)) > mLateAllow Then
                                                            .Fields("latehrs") = DiffHrs(CDate(mT2in), CDate(mST1out)) - mLateAllow
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'compute for undertime
                                            If CDate(mT2out) < CDate(mST2out) Then
                                                If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                    .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !firsttravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mST1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                                mNiteWrkHrs = mNiteWrkHrs + DiffNiteHrs(mT2in, mT2out, mST2out, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in <> "" And mT2in = "" Then
                                        
                                            If CDate(mT1in) <= CDate(mST2in) Then
                                                'compute for undertime
                                                If CDate(mT1out) > (mST2in) And CDate(mT1out) < CDate(mST2out) Then
                                                    If DiffHrs(CDate(mT1out), CDate(mST2out)) > 0 Then
                                                        .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST2out))
                                                    End If
                                                    mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                                    If !firsttravel = 1 Then
                                                        .Fields("dayswrk") = 1
                                                        mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                ElseIf CDate(mT1out) > CDate(mST2out) Then
                                                    mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                                    If !firsttravel = 1 Then
                                                        .Fields("dayswrk") = 1
                                                        mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                ElseIf CDate(mT1out) <= CDate(mST2in) Then
                                                    .Fields("absent") = 0.5
                                                    .Fields("daywrk") = 0.5
                                                    If !firsttravel = 1 Then
                                                        mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                                    End If
                                                End If
                                            End If
                                            If !firsttravel = 1 Then
                                                mNiteWrkHrs = DiffNiteHrs(mST1in, mT1out, mST1in, mST2out, !nitepremstart, !nitepremend)
                                            Else
                                                mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            End If
                                        ElseIf mT1in = "" And mT2in <> "" Then
                                            'compute for late
                                            If CDate(mT2in) > CDate(mST2in) Then
                                                If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                    If DiffHrs(CDate(mST2in), CDate(mT2in)) > mLateAllow Then
                                                        .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in)) - mLateAllow
                                                    End If
                                                End If
                                            End If
                                            
                                            'compute for undertime
                                            If CDate(mT2out) < CDate(mST2out) Then
                                                If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                                    .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                                End If
                                            End If
                                            mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out)) - .Fields("uthrs") - .Fields("latehrs")
                                            If !firsttravel = 1 Then
                                                .Fields("dayswrk") = 1
                                                mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                                mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                            Else
                                                .Fields("dayswrk") = 0.5
                                                .Fields("absent") = 0.5
                                            End If
                                        End If
                                    Else
                                        .Fields("dayswrk") = 1
                                        If !firsttravel = 1 Then
                                            mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                            mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                        End If
                                        If !secondtravel = 1 Then
                                            mNiteWrkHrs = DiffNiteHrs(mT2in, mT2out, mST2in, mST2out, !nitepremstart, !nitepremend)
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                        End If
                                    End If
                                End If
                            Else 'For schedules with only two(2) time slots
                                If !travel = 0 And !leave = 0 Then
                                    If mT1in <> "" Then 'Check if time slots were used.
                                        .Fields("dayswrk") = 1
                                        If CDate(mT1in) < CDate(mST1in) Then
                                            mTin = mST1in
                                        Else
                                            mTin = mT1in
                                            'compute for late
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                If DiffHrs(CDate(mST1in), CDate(mT1in)) > mLateAllow Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in)) - mLateAllow
                                                End If
                                            End If
                                        End If
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            mTout = mT1out
                                            'compute for undertime
                                            If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                                .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                            End If
                                        Else
                                            mTout = mST1out
                                        End If
                                        mNiteWrkHrs = DiffNiteHrs(mT1in, mT1out, mST1in, mST1out, !nitepremstart, !nitepremend)
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                        
                                        'mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    Else 'absent
                                        If Trim(.Fields("holiday")) = "" Then
                                            If .Fields("required") = "Y" Then
                                                .Fields("absent") = 1
                                            End If
                                            .Fields("dayswrk") = 0
                                        End If
                                    End If
                                End If
                            End If
                            .Fields("wrkhrs") = mWrkHrs - mNiteWrkHrs
                            .Fields("nitewrkhrs") = mNiteWrkHrs
                        End If
                        
                        If Trim(!hrsperday) = "" Then !hrsperday = 0
                        If Trim(!brkhrsperday) = "" Then !brkhrsperday = 0

                        ConMain.Execute "delete from dtremp where employeecode = '" & rsEmployee!employeecode & "' and workdate ='" & Format(!wrkdate, "YYYY-MM-DD") & "'"
                        
                        ConMain.Execute "insert into dtremp(employeecode,payfreqcode,dayno,workdate,shiftcode, " & _
                                              "t1in,t1out,t2in,t2out,st1in,st1out,st2in,st2out, " & _
                                              "wrkhrs,nitewrkhrs,dayswrk,absent,latehrs,uthrs,dayoff,updatable,brkstart,brkend, " & _
                                              "nitepremstart,nitepremend,hrsperday,brkhrsperday,required, " & _
                                              "branchcode,divisioncode,costcentercode,sectioncode, holiday, " & _
                                              "firstleave,secondleave) values " & _
                                              "('" & rsEmployee!employeecode & "','" & tdbPayrollPeriod.Columns("payfreqcode").Text & "','" & !dayno & "','" & Format(!wrkdate, "YYYY-MM-DD") & "','" & !shiftcode & "', " & _
                                              "'" & Format(!t1in, "hh:nn") & "','" & Format(!t1out, "hh:nn") & "','" & Format(!t2in, "hh:nn") & "','" & Format(!t2out, "hh:nn") & "', " & _
                                              "'" & !st1in & "','" & !st1out & "','" & !st2in & "','" & !st2out & "', " & _
                                              !wrkhrs & "," & !nitewrkhrs & "," & !dayswrk & "," & !absent & "," & !latehrs & "," & !uthrs & ",'" & IIf(!dayoff <> 0, "Y", "N") & "','" & IIf(!updatable <> 0, "Y", "N") & "','" & !brkstart & "','" & !brkend & "', " & _
                                              "'" & !nitepremstart & "','" & !nitepremend & "'," & !hrsperday & "," & !brkhrsperday & ",'" & !Required & "', " & _
                                              "'" & rsEmployee!branchcode & "','" & rsEmployee!divisioncode & "','" & rsEmployee!costcentercode & "','" & rsEmployee!sectioncode & "','" & !Holiday & "', " & _
                                              "'" & !firstleave & "','" & !secondleave & "')"

                        .MoveNext
                        DoEvents
                        
                    Loop
                    
                    .MoveFirst
                    
                End With
                
            End If
            
            'Computes Overtime
            
            Set rsOTTmp = Nothing
            Set rsOTTmp = New ADODB.Recordset
            
            With rsOTTmp
                .Fields.Append "otlneno", adVarChar, 7
                .Fields.Append "otcode", adVarChar, 15
                .Fields.Append "employeecode", adVarChar, 15
                .Fields.Append "percode", adVarChar, 7
                .Fields.Append "dayoff", adInteger
                .Fields.Append "holiday", adVarChar, 10
                .Fields.Append "wrkdate", adDate
                .Fields.Append "day", adVarChar, 15
                .Fields.Append "actotstart", adVarChar, 8
                .Fields.Append "actotend", adVarChar, 8
                .Fields.Append "otstart", adVarChar, 5
                .Fields.Append "otend", adVarChar, 5
                .Fields.Append "otwrkhrs", adDouble, 18
                .Fields.Append "regwrkhrs", adDouble, 18
                .Fields.Append "nitewrkhrs", adDouble, 18
                .Open
            End With
            
            NetOpen rsOT, "select * from overtimelne where employeecode = '" & rsEmployee!employeecode & "' and status = 'Approved' and percode = '" & tdbPayrollPeriod.Columns("percode").Text & "' and fnlz <> 'Y' order by wrkdate,otstart"
            
            With rsOTTmp
                If rsOT.RecordCount > 0 Then
                    rsOT.MoveFirst
                    
                    pb2.Value = 0
                    pb2.Max = rsOT.RecordCount
                    Do While Not rsOT.EOF
                    
                        Set rsOTCmpr = New ADODB.Recordset
                        Set rsOTCmpr = rsOTTmp.Clone
                        
                        pb2.Value = pb2.Value + 1
                        
                        mOT_Start = ""
                        mOT_End = ""
                        mNiteStart = ""
                        mNiteEnd = ""
                        mTimeIN = ""
                        mTimeOUT = ""
                        
                        .AddNew
                        .Fields("otlneno") = rsOT!otlneno
                        .Fields("otcode") = rsOT!otcode
                        .Fields("employeecode") = rsOT!employeecode
                        .Fields("percode") = rsOT!percode
                        .Fields("wrkdate") = rsOT!wrkdate
                        .Fields("day") = WeekdayName(Weekday(rsOT!wrkdate))
                        .Fields("otstart") = rsOT!otstart
                        .Fields("otend") = rsOT!otend
                        
                        'Check if dayoff
                        NetOpen rsDtrEmp, "select * from dtremp where employeecode = '" & rsEmployee!employeecode & "' and workdate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "' "
                        
                        If rsDtrEmp.RecordCount > 0 Then
                            .Fields("dayoff") = IIf(rsDtrEmp!dayoff = "Y", 1, 0)
                            'Check if it has a night premium hours
                            If Trim(rsDtrEmp!nitepremstart) <> "" Then
                                mNiteStart = Format(rsOT!wrkdate & " " & rsDtrEmp!nitepremstart, "MM/DD/YYYY hh:nn:ss")
                                If CDate(mNiteStart) > Format(rsOT!wrkdate & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss") Then
                                    mNiteEnd = Format(CDate(rsOT!wrkdate) + 1 & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss")
                                Else
                                    mNiteEnd = Format(CDate(rsOT!wrkdate) & " " & rsDtrEmp!nitepremend, "MM/DD/YYYY hh:nn:ss")
                                End If
                            End If
                        End If
                        
                        'Check if Holiday
                        NetOpen rsHoliday, "select * from holiday where holidaydate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
                        If rsHoliday.RecordCount > 0 Then
                            If CInt(rsHoliday!regular) = 1 Then
                              .Fields("holiday") = "Legal"
                            Else
                              .Fields("holiday") = "Special"
                            End If
                        Else
                            .Fields("holiday") = ""
                        End If
                        
                        Create_TmpOTDTSTito (rsOT!wrkdate)
                        
                        mOT_Start = Format(rsOT!wrkdate & " " & rsOT!otstart, "MM/DD/YYYY hh:nn:ss")
                        
                        'Check if Overtime out is on the next day
                        
                        If CDate(mOT_Start) > Format(rsOT!wrkdate & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss") Then
                            mOT_End = Format(CDate(rsOT!wrkdate) + 1 & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss")
                        Else
                            mOT_End = Format(rsOT!wrkdate & " " & rsOT!otend, "MM/DD/YYYY hh:nn:ss")
                        End If
                        
                        If rsTitoTmp.RecordCount > 0 Then
                            
                            rsTitoTmp.Filter = "tout > '" & mOT_Start & "' and tin < '" & mOT_End & "'"
                            
                            If Not rsTitoTmp.EOF Then
                            
                                rsTitoTmp.MoveFirst
                                
                                !otwrkhrs = 0
                                !regwrkhrs = 0
                                !nitewrkhrs = 0
                                
                                .Fields("actotstart") = Format(rsTitoTmp!tin, "hh:nn:ss")
                                
                                Do While Not rsTitoTmp.EOF
                                
                                    .Fields("actotend") = Format(rsTitoTmp!tout, "hh:nn:ss")
                                    
                                    If CDate(rsTitoTmp!tin) < CDate(mOT_Start) Then
                                        mOTActIN = mOT_Start
                                    Else
                                        mOTActIN = rsTitoTmp!tin
                                    End If
                                    
                                    If CDate(rsTitoTmp!tout) > CDate(mOT_End) Then
                                        mOTActOUT = mOT_End
                                    Else
                                        mOTActOUT = rsTitoTmp!tout
                                    End If
                                    
                                    !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                                              
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                            
                                        End If
                                    End If
                                    
                                    rsTitoTmp.MoveNext
                                    
                                Loop
                            End If
                        End If
                        
                        'Computes manual TITO using Employee's DTR
                        
                        'only if Employee's TITO is not available from TITO TABLE
                        
                        If Trim(!actotstart) = "" Then
                            
                            !otwrkhrs = 0
                            !regwrkhrs = 0
                            !nitewrkhrs = 0
                            
                            If rsDtrEmp.RecordCount > 0 Then
                            
                                If Trim(rsDtrEmp!t1in) <> "" Then
                                    
                                    .Fields("actotstart") = rsDtrEmp!t1in
                                    .Fields("actotend") = rsDtrEmp!t1out
                                    
                                    mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t1in, "MM/DD/YYYY hh:nn:ss")
                                    If CDate(mTimeIN) > Format(rsOT!wrkdate & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss") Then
                                        mTimeOUT = Format(CDate(rsOT!wrkdate) + 1 & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss")
                                    Else
                                        mTimeOUT = Format(rsOT!wrkdate & " " & rsDtrEmp!t1out, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeOUT) > CDate(mOT_Start) And CDate(mTimeIN) < CDate(mOT_End) Then
                                        If CDate(mTimeIN) < CDate(mOT_Start) Then
                                            mOTActIN = mOT_Start
                                        Else
                                            mOTActIN = mTimeIN
                                        End If
                                        
                                        If CDate(mTimeOUT) > CDate(mOT_End) Then
                                            mOTActOUT = mOT_End
                                        Else
                                            mOTActOUT = mTimeOUT
                                        End If
                                        !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                    End If
                                    
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                        End If
                                    End If
                                    
                                End If
                                                                                
                                If Trim(rsDtrEmp!t2in) <> "" Then
                                
                                    If Trim(rsDtrEmp!t1in) = "" Then
                                        .Fields("actotstart") = rsDtrEmp!t2in
                                    End If
                                    .Fields("actotend") = rsDtrEmp!t2out
                                    
                                    If Trim(mTimeOUT) <> "" Then
                                        If CDate(mTimeOUT) > Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss") Then
                                            mTimeIN = Format(CDate(Format(mTimeOUT, "MM/DD/YYYY")) + 1 & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                        Else
                                            mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                        End If
                                    Else
                                        mTimeIN = Format(rsOT!wrkdate & " " & rsDtrEmp!t2in, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeIN) > Format(rsOT!wrkdate & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss") Then
                                        mTimeOUT = Format(CDate(Format(mTimeIN, "MM/DD/YYYY")) + 1 & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss")
                                    Else
                                        mTimeOUT = Format(rsOT!wrkdate & " " & rsDtrEmp!t2out, "MM/DD/YYYY hh:nn:ss")
                                    End If
                                    
                                    If CDate(mTimeOUT) > CDate(mOT_Start) And CDate(mTimeIN) < CDate(mOT_End) Then
                                        If CDate(mTimeIN) < CDate(mOT_Start) Then
                                            mOTActIN = mOT_Start
                                        Else
                                            mOTActIN = mTimeIN
                                        End If
                                        
                                        If CDate(mTimeOUT) > CDate(mOT_End) Then
                                            mOTActOUT = mOT_End
                                        Else
                                            mOTActOUT = mTimeOUT
                                        End If
                                        !otwrkhrs = !otwrkhrs + DiffHrs(mOTActIN, mOTActOUT)
                                    End If
                                    
                                    'Compute night premium hours if any
                                    'Check if time falls on night shift premium
                                    If Trim(mNiteStart) <> "" Then
                                        If mOTActOUT > CDate(mNiteStart) And mOTActIN < CDate(mNiteEnd) Then
                                            If mOTActIN < CDate(mNiteStart) Then
                                                mNiteIN = mNiteStart
                                            Else
                                                mNiteIN = mOTActIN
                                            End If
                                            If mOTActOUT > CDate(mNiteEnd) Then
                                                mNiteOUT = mNiteEnd
                                            Else
                                                mNiteOUT = mOTActOUT
                                            End If
                                            !nitewrkhrs = !nitewrkhrs + DiffHrs(mNiteIN, mNiteOUT)
                                        End If
                                        
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        If !otwrkhrs >= mOTAllow Then
                            !otwrkhrs = !otwrkhrs - !nitewrkhrs
                        Else
                            !otwrkhrs = 0
                            !nitewrkhrs = 0
                        End If
                        
                        mLackHrs = 0
                        mTtlReghrs = 0
                        
                        If !otwrkhrs > 0 Then
                            NetOpen rsOTChk, "select (hrsperday - wrkhrs) lackhrs,hrsperday,firstleave,secondleave from dtremp where hrsperday > wrkhrs and employeecode = '" & rsEmployee!employeecode & "' and workdate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
                            If rsOTChk.RecordCount > 0 Then
                                
                                If rsOTChk!lackhrs >= mUTMin Then
                                
                                    mLackHrs = rsOTChk!lackhrs
                                    
                                    If rsOTChk!firstleave = 1 Then
                                        mLackHrs = mLackHrs - Format(rsOTChk!hrsperday / 2, "#,##0.00")
                                    End If
                                    
                                    If rsOTChk!secondleave = 1 Then
                                        mLackHrs = mLackHrs - Format(rsOTChk!hrsperday / 2, "#,##0.00")
                                    End If
                                    
                                    If mLackHrs > 0 Then
                                        
                                        If rsOTCmpr.RecordCount > 0 Then
                                            rsOTCmpr.MoveFirst
                                            Do While Not rsOTCmpr.EOF
                                                If rsOTCmpr!otlneno <> !otlneno And rsOTCmpr!wrkdate = rsOT!wrkdate Then
                                                    mTtlReghrs = mTtlReghrs + rsOTCmpr!regwrkhrs
                                                End If
                                                rsOTCmpr.MoveNext
                                            Loop
                                        End If
                                        
                                        If mLackHrs > mTtlReghrs Then
                                            mLackHrs = mLackHrs - mTtlReghrs
                                            If mLackHrs > !otwrkhrs Then
                                                !regwrkhrs = !otwrkhrs
                                                !otwrkhrs = 0
                                            Else
                                                !regwrkhrs = mLackHrs
                                                !otwrkhrs = !otwrkhrs - mLackHrs
                                            End If
                                        
                                        End If
                                    
                                    End If
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                        .Update
                        
                        ConMain.Execute "update overtimelne set actotstart = '" & !actotstart & "', actotend = '" & !actotend & "',otwrkhrs = " & !otwrkhrs & ", regwrkhrs = " & !regwrkhrs & ",nitewrkhrs = " & !nitewrkhrs & ", holiday = '" & !Holiday & "',dayoff = '" & IIf(!dayoff <> 0, "Y", "N") & "' " & _
                                              "where otlneno = '" & !otlneno & "'"
                        
                        rsOT.MoveNext
                    Loop
                    
                End If
                                
            End With
            
            rsEmployee.MoveNext
            
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

Private Sub Create_TmpTito()
    
    Dim rsTito      As ADODB.Recordset
    Dim mDateTmp    As String
    Dim isIn        As Boolean

    Set rsTitoTmp = Nothing
    Set rsTitoTmp = New ADODB.Recordset

    With rsTitoTmp
        .Fields.Append "wrkdate", adDate
        .Fields.Append "tin", adDate
        .Fields.Append "tout", adDate
        .Open
        .Sort = "wrkdate"
    End With

    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
                        "from tito where employeecode = '" & rsEmployee!employeecode & "' and " & _
                        "datelog Between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "'  and " & _
                        "'" & Format(CDate(tdbPayrollPeriod.Columns("wrkdateto").Text) + 1, "YYYY-MM-DD") & "' " & _
                        "Union All " & _
                        "select employeecode,complog,datelog,timelog,logstat " & _
                        "from gplne where employeecode = '" & rsEmployee!employeecode & "' and " & _
                        "(datelog Between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' and " & _
                        "'" & Format(CDate(tdbPayrollPeriod.Columns("wrkdateto").Text) + 1, "YYYY-MM-DD") & "') and  " & _
                        "percode = '" & tdbPayrollPeriod.BoundText & "' and status = 'Approved' " & _
                        "order by complog"
                        
    With rsTito
        If .RecordCount > 0 Then

            .MoveFirst
            pb2.Max = .RecordCount
            pb2.Value = 0
            If !logstat = "Out" Then
                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
            End If
            isIn = False
            
            Do While Not .EOF
                
                pb2.Value = pb2.Value + 1
                
                If !logstat = "In" Then
                
                    If Not isIn Then
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    Else
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    End If
                    
                Else
                    If isIn Then
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    Else
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = Format(mDateTmp, "MM/DD/YYYY")
                        rsTitoTmp.Fields("tin") = mDateTmp
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    End If
                End If
                .MoveNext
                DoEvents
            Loop

        End If

    End With
    
End Sub

Private Sub Create_TmpOTDTSTito(mDate As Date)
    
    Dim rsTito      As ADODB.Recordset
    Dim mDateTmp    As String
    Dim isIn        As Boolean

    Set rsTitoTmp = Nothing
    Set rsTitoTmp = New ADODB.Recordset

    With rsTitoTmp
        .Fields.Append "wrkdate", adDate
        .Fields.Append "tin", adDate
        .Fields.Append "tout", adDate
        .Open
        .Sort = "wrkdate"
    End With


    NetOpen rsTito, "select employeecode,complog,datelog,timelog,logstat " & _
                        "from tito where employeecode = '" & rsEmployee!employeecode & "' and " & _
                        "datelog Between '" & Format(mDate - 1, "YYYY-MM-DD") & "'  and " & _
                        "'" & Format(mDate + 1, "YYYY-MM-DD") & "' " & _
                        "Union All " & _
                        "select employeecode,complog,datelog,timelog,logstat " & _
                        "from gplne where employeecode = '" & rsEmployee!employeecode & "' and " & _
                        "(datelog Between '" & Format(mDate - 1, "YYYY-MM-DD") & "' and " & _
                        "'" & Format(mDate + 1, "YYYY-MM-DD") & "') and  " & _
                        "percode = '" & tdbPayrollPeriod.BoundText & "' and status = 'Approved' " & _
                        "order by complog"
    
    With rsTito
        If .RecordCount > 0 Then

            .MoveFirst
            
            If !logstat = "Out" Then
                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
            End If
            isIn = False
            
            Do While Not .EOF
                If !logstat = "In" Then
                
                    If Not isIn Then
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    Else
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = !datelog
                        rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        isIn = True
                    End If
                    
                Else
                
                    If isIn Then
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    Else
                        rsTitoTmp.AddNew
                        rsTitoTmp.Fields("wrkdate") = Format(mDateTmp, "MM/DD/YYYY")
                        rsTitoTmp.Fields("tin") = mDateTmp
                        rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                        rsTitoTmp.Update
                        isIn = False
                        mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                    End If
                    
                End If
                .MoveNext
                DoEvents
            Loop
            rsTitoTmp.MoveLast
            If Trim(rsTitoTmp!tout) = "" Then
                rsTitoTmp.Delete
            End If

        End If

    End With
    
End Sub


Private Sub Load_Dtr()

    Dim mDate       As Date
    
    Dim rsEmpDtr    As ADODB.Recordset
    Dim rsEmpShift  As ADODB.Recordset
    Dim rsEmpShift2 As ADODB.Recordset
    Dim rsHoliday   As ADODB.Recordset
    Dim rsOBT       As ADODB.Recordset
    Dim rsLeave     As ADODB.Recordset
    
    Set rsDtrTmp = Nothing
    Set rsDtrTmp = New ADODB.Recordset
    
    With rsDtrTmp
  
        .Fields.Append "updatable", adInteger
        .Fields.Append "wrkdate", adDate
        .Fields.Append "dayno", adInteger
        .Fields.Append "day", adVarChar, 15
        .Fields.Append "dayoff", adInteger
        .Fields.Append "dayswrk", adDouble, 18
        .Fields.Append "holiday", adVarChar, 10
        .Fields.Append "Travel", adInteger
        .Fields.Append "Leave", adInteger
        .Fields.Append "t1in", adVarChar, 15, adFldIsNullable
        .Fields.Append "t1out", adVarChar, 15, adFldIsNullable
        .Fields.Append "t2in", adVarChar, 15, adFldIsNullable
        .Fields.Append "t2out", adVarChar, 15, adFldIsNullable
        .Fields.Append "st1in", adVarChar, 5
        .Fields.Append "st1out", adVarChar, 5
        .Fields.Append "st2in", adVarChar, 5
        .Fields.Append "st2out", adVarChar, 5
        .Fields.Append "brkstart", adVarChar, 5
        .Fields.Append "brkend", adVarChar, 5
        .Fields.Append "nitepremstart", adVarChar, 5
        .Fields.Append "nitepremend", adVarChar, 5
        .Fields.Append "shiftcode", adVarChar, 7
        .Fields.Append "shiftdetail", adVarChar, 50
        .Fields.Append "wrkhrs", adDouble
        .Fields.Append "nitewrkhrs", adDouble
        .Fields.Append "absent", adDouble
        .Fields.Append "latehrs", adDouble
        .Fields.Append "uthrs", adDouble
        .Fields.Append "hrsperday", adDouble, 18
        .Fields.Append "brkhrsperday", adDouble, 18
        .Fields.Append "firsttravel", adVarChar, 1
        .Fields.Append "secondtravel", adVarChar, 1
        .Fields.Append "firstleave", adVarChar, 1
        .Fields.Append "secondleave", adVarChar, 1
        .Fields.Append "required", adVarChar, 1
        .Open

        mDate = Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "MM/DD/YYYY")
        pb2.Max = DateDiff("d", mDate, Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")) + 1
        pb2.Value = 0
        
        Do While mDate <= Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")
            
            pb2.Value = pb2.Value + 1
            .AddNew
            .Fields("wrkdate") = mDate
            .Fields("dayno") = Weekday(mDate)
            .Fields("day") = WeekdayName(Weekday(mDate))
            
            NetOpen rsEmpShift, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from empshift x1 left outer join shift x2 on " & _
                                  "x1.shiftcode = x2.shiftcode where x1.shiftcode <> '' and x1.employeecode = '" & rsEmployee!employeecode & "' and x1.dayno = '" & Weekday(mDate) & "'"
            
            If rsEmpShift.RecordCount > 0 Then
        
                .Fields("updatable") = 1
                .Fields("t1in") = ""
                .Fields("t1out") = ""
                .Fields("t2in") = ""
                .Fields("t2out") = ""
                .Fields("st1in") = rsEmpShift!t1in
                .Fields("st1out") = rsEmpShift!t1out
                .Fields("st2in") = rsEmpShift!t2in
                .Fields("st2out") = rsEmpShift!t2out
                .Fields("shiftcode") = rsEmpShift!shiftcode
                .Fields("shiftdetail") = rsEmpShift!t1in & "   " & rsEmpShift!t1out & "       " & rsEmpShift!t2in & "   " & rsEmpShift!t2out
                .Fields("brkstart") = rsEmpShift!brkstart
                .Fields("brkend") = rsEmpShift!brkend
                .Fields("nitepremstart") = rsEmpShift!nitepremstart
                .Fields("nitepremend") = rsEmpShift!nitepremend
                .Fields("hrsperday") = rsEmpShift!hrsperday
                .Fields("brkhrsperday") = rsEmpShift!brkhrsperday
                .Fields("dayoff") = 0
                .Fields("required") = rsEmpShift!Required
                
                
                NetOpen rsEmpDtr, "select * from dtremp where employeecode = '" & rsEmployee!employeecode & "' and " & _
                                    "workdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
                                      
                If rsEmpDtr.RecordCount > 0 Then
                    NetOpen rsEmpShift2, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from  shift x2 where x2.shiftcode = '" & rsEmpDtr!shiftcode & "'"
                    If rsEmpShift2.RecordCount > 0 Then
                        If rsEmpDtr!updatable = "N" Then
                            .Fields("updatable") = 0
                            .Fields("t1in") = rsEmpDtr!t1in
                            .Fields("t1out") = rsEmpDtr!t1out
                            .Fields("t2in") = rsEmpDtr!t2in
                            .Fields("t2out") = rsEmpDtr!t2out
                            
                            .Fields("st1in") = rsEmpShift2!t1in
                            .Fields("st1out") = rsEmpShift2!t1out
                            .Fields("st2in") = rsEmpShift2!t2in
                            .Fields("st2out") = rsEmpShift2!t2out
                            
                        End If
                        
                        .Fields("shiftcode") = rsEmpShift2!shiftcode
                        .Fields("shiftdetail") = rsEmpShift2!t1in & "   " & rsEmpShift2!t1out & "       " & rsEmpShift2!t2in & "   " & rsEmpShift2!t2out
                        .Fields("brkstart") = rsEmpShift2!brkstart
                        .Fields("brkend") = rsEmpShift2!brkend
                        .Fields("nitepremstart") = rsEmpShift2!nitepremstart
                        .Fields("nitepremend") = rsEmpShift2!nitepremend
                        .Fields("hrsperday") = rsEmpShift2!hrsperday
                        .Fields("brkhrsperday") = rsEmpShift2!brkhrsperday
                        .Fields("required") = rsEmpShift2!Required
                    Else
                        .Fields("st1in") = ""
                        .Fields("st1out") = ""
                        .Fields("st2in") = ""
                        .Fields("st2out") = ""
                        .Fields("shiftcode") = ""
                        .Fields("shiftdetail") = ""
                        .Fields("brkstart") = ""
                        .Fields("brkend") = ""
                        .Fields("nitepremstart") = ""
                        .Fields("nitepremend") = ""
                        .Fields("hrsperday") = 0
                        .Fields("brkhrsperday") = 0
                        .Fields("dayoff") = 1
                        .Fields("required") = "N"
                    End If
                End If
            Else
            
                NetOpen rsEmpDtr, "select * from dtremp where employeecode = '" & rsEmployee!employeecode & "' and " & _
                                    "workdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
                                      
                If rsEmpDtr.RecordCount > 0 Then
                    NetOpen rsEmpShift2, "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from  shift x2 where x2.shiftcode = '" & rsEmpDtr!shiftcode & "'"
                    If rsEmpShift2.RecordCount > 0 Then
                        If rsEmpDtr!updatable = "N" Then
                            .Fields("updatable") = 0
                            .Fields("t1in") = rsEmpDtr!t1in
                            .Fields("t1out") = rsEmpDtr!t1out
                            .Fields("t2in") = rsEmpDtr!t2in
                            .Fields("t2out") = rsEmpDtr!t2out
                            
                            .Fields("st1in") = rsEmpShift2!t1in
                            .Fields("st1out") = rsEmpShift2!t1out
                            .Fields("st2in") = rsEmpShift2!t2in
                            .Fields("st2out") = rsEmpShift2!t2out
                            
                        End If
                        
                        .Fields("shiftcode") = rsEmpShift2!shiftcode
                        .Fields("shiftdetail") = rsEmpShift2!t1in & "   " & rsEmpShift2!t1out & "       " & rsEmpShift2!t2in & "   " & rsEmpShift2!t2out
                        .Fields("brkstart") = rsEmpShift2!brkstart
                        .Fields("brkend") = rsEmpShift2!brkend
                        .Fields("nitepremstart") = rsEmpShift2!nitepremstart
                        .Fields("nitepremend") = rsEmpShift2!nitepremend
                        .Fields("hrsperday") = rsEmpShift2!hrsperday
                        .Fields("brkhrsperday") = rsEmpShift2!brkhrsperday
                        .Fields("required") = rsEmpShift2!Required
                        .Fields("dayoff") = IIf(rsEmpDtr!dayoff = "Y", 1, 0)
                    Else
                        .Fields("st1in") = ""
                        .Fields("st1out") = ""
                        .Fields("st2in") = ""
                        .Fields("st2out") = ""
                        .Fields("shiftcode") = ""
                        .Fields("shiftdetail") = ""
                        .Fields("brkstart") = ""
                        .Fields("brkend") = ""
                        .Fields("nitepremstart") = ""
                        .Fields("nitepremend") = ""
                        .Fields("hrsperday") = 0
                        .Fields("brkhrsperday") = 0
                        .Fields("dayoff") = 1
                        .Fields("required") = "N"
                    End If
                Else
                    .Fields("updatable") = 1
                    .Fields("dayoff") = 1
                    .Fields("required") = "N"
                End If
                  
            End If
              
            NetOpen rsHoliday, "select x1.* from holiday x1 " & _
                    "left outer join holidaybranchinclude x2 on x1.holidaydate = x2.holidaydate " & _
                    "where x1.holidaydate = '" & Format(mDate, "YYYY-MM-DD") & "' and x2.branchcode = '" & rsEmployee!branchcode & "'"
            
            If rsHoliday.RecordCount > 0 Then
                If CInt(rsHoliday!regular) = 1 Then
                  .Fields("holiday") = "Legal"
                Else
                  .Fields("holiday") = "Special"
                End If
            Else
                .Fields("holiday") = ""
            End If
            
            NetOpen rsOBT, "select x1.* from obtlne x1 left outer join obthdr x2 on x1.obtnum = x2.obtnum " & _
                             "where x2.employeecode = '" & rsEmployee!employeecode & "' and x1.obtdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
            
            If rsOBT.RecordCount > 0 Then
                .Fields("travel") = 1
                .Fields("firsttravel") = rsOBT!firstshift
                .Fields("secondtravel") = rsOBT!secondshift
            Else
                .Fields("travel") = 0
                .Fields("firsttravel") = 0
                .Fields("secondtravel") = 0
            End If
            
            NetOpen rsLeave, "select x1.* from lvlne x1 " & _
                    "left outer join lvhdr x2 on x1.lvnum = x2.lvnum " & _
                    "where x2.employeecode = '" & rsEmployee!employeecode & "' and x1.lvdate = '" & Format(mDate, "YYYY-MM-DD") & "'"
            
            If rsLeave.RecordCount > 0 Then
                .Fields("leave") = 1
                .Fields("firstleave") = rsLeave!firstshift
                .Fields("secondleave") = rsLeave!secondshift
            Else
                .Fields("leave") = 0
                .Fields("firstleave") = 0
                .Fields("secondleave") = 0
            End If
            
            .Update
            
            mDate = mDate + 1
            DoEvents
        Loop
      
    End With

End Sub

Private Function DiffHrs(mHrs1 As Date, mHrs2 As Date) As Double
    DiffHrs = Format(Round(DateDiff("N", mHrs1, mHrs2) / 60, 2), "#,##0.00")
End Function


Private Function DiffNiteHrs(ByVal mActIn As Variant, ByVal mActOut As Variant, ByVal mSTin As String, ByVal mSTout As String, ByVal mNiteStart As String, ByVal mNiteEnd As String) As Double
    
    Dim mNiteIN     As Date
    Dim mNiteOUT    As Date
    Dim mDumIn      As Date
    Dim mDumOut     As Date
    
    DiffNiteHrs = 0
    
    If Trim(mNiteStart) <> "" Then
        If CDate(mActIn) < CDate(mSTin) Then
            mDumIn = CDate(mSTin)
        Else
            mDumIn = CDate(mActIn)
        End If
        If CDate(mActOut) > CDate(mSTout) Then
            If mDumIn > CDate(mSTout) Then
                mDumOut = mDumIn
            Else
                mDumOut = CDate(mSTout)
            End If
        Else
            mDumOut = CDate(mActOut)
        End If
        If mDumOut > CDate(mNiteStart) And mDumIn < CDate(mNiteEnd) Then
            If mDumIn < CDate(mNiteStart) Then
                mNiteIN = mNiteStart
            Else
                mNiteIN = mDumIn
            End If
            If mDumOut > CDate(mNiteEnd) Then
                mNiteOUT = mNiteEnd
            Else
                mNiteOUT = mDumOut
            End If
            DiffNiteHrs = DiffHrs(mNiteIN, mNiteOUT)
        End If
    End If
    
End Function

Private Sub tdbBranch_ItemChange()

    bind_tdb ConMain, tdbDivision, "select divisioncode,division from division " & _
              "where branchcode = '" & tdbBranch.BoundText & "' order by division", "division", "divisioncode"
    
    tdbDivision.BoundText = ""
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
  
  bind_tdb ConMain, tdbCostCenter, "select costcentercode,costcenter from costcenter " & _
            "where branchcode = '" & tdbBranch.BoundText & "' order by costcenter", "costcenter", "costcentercode"

    tdbCostCenter.BoundText = ""
  
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

    bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname, ', ', firstname,' ',middlename) fullname from employee " & _
                "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and costcentercode = '" & tdbCostCenter.BoundText & "' " & _
                "and payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' " & _
                "order by concat(lastname, ', ', firstname,' ',middlename) ", "fullname", "employeecode"
    
    tdbEmployee.BoundText = ""
    
End Sub

Private Sub tdbcostcenter_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchList KeyAscii, tdbCostCenter, tdbCostCenter.RowSource, tdbCostCenter.Text
      tdbcostcenter_ItemChange
    End If
    
End Sub

Private Sub tdbEmployee_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbEmployee, tdbEmployee.RowSource, tdbEmployee.Text
    End If
End Sub

Private Sub tdbPayrollPeriod_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
    End If
End Sub
