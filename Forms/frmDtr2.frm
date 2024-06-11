VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDtr2 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   315
      TabIndex        =   1
      Top             =   1605
      Width           =   12720
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import TITO"
         Height          =   360
         Left            =   8280
         TabIndex        =   8
         Top             =   1080
         Width           =   2025
      End
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   345
         Left            =   1530
         TabIndex        =   2
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
         _PropDict       =   $"frmDtr2.frx":0000
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
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1530
         TabIndex        =   3
         Tag             =   "Municipal"
         Top             =   255
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
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         _PropDict       =   $"frmDtr2.frx":00AA
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   90
         Left            =   8280
         TabIndex        =   7
         Top             =   1455
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   159
         _Version        =   393216
         Appearance      =   0
      End
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   330
         Left            =   5835
         TabIndex        =   10
         Top             =   300
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   582
         Caption         =   "&Generate"
         CapAlign        =   2
         BackStyle       =   4
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
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
         Height          =   225
         Left            =   255
         TabIndex        =   5
         Top             =   285
         Width           =   1830
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         Height          =   225
         Left            =   255
         TabIndex        =   4
         Top             =   675
         Width           =   1830
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgTito 
      Height          =   3060
      Left            =   210
      TabIndex        =   0
      Top             =   2775
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   5398
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Date"
      Columns(0).DataField=   "Wrkdate"
      Columns(0).NumberFormat=   "MM/DD/YYYY"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Time IN"
      Columns(1).DataField=   "tin"
      Columns(1).NumberFormat=   "MM/DD/YYYY HH:NN:SS AM/PM"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Time OUT"
      Columns(2).DataField=   "tout"
      Columns(2).NumberFormat=   "MM/DD/YYYY HH:NN:SS AM/PM"
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
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3387"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3307"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1693"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1614"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   16185592
      RowDividerColor =   16185592
      RowSubDividerColor=   16185592
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF6F8F8&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H400000&"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HF6F8F8&"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H0&,.fgcolor=&H0&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HEBFEEB&,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H0&"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(17)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFF0EA&"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HF6F8F8&"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid tdgDtr 
      Height          =   3210
      Left            =   5745
      TabIndex        =   6
      Top             =   2685
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   5662
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Date"
      Columns(0).DataField=   "wrkdate"
      Columns(0).NumberFormat=   "MM/DD/YYYY"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Day"
      Columns(1).DataField=   "day"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Holiday"
      Columns(2).DataField=   "holiday"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   4
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "On Travel"
      Columns(3).DataField=   "travel"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "On Leave"
      Columns(4).DataField=   "leave"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Day Off"
      Columns(5).DataField=   "dayoff"
      Columns(5).DataWidth=   1
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "1st Time In"
      Columns(6).DataField=   "t1in"
      Columns(6).NumberFormat=   "hh:nn:ss"
      Columns(6).ExternalEditor=   "txtTime"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "1st Time Out"
      Columns(7).DataField=   "t1Out"
      Columns(7).NumberFormat=   "hh:nn:ss"
      Columns(7).ExternalEditor=   "txtTime"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "2nd Time In"
      Columns(8).DataField=   "t2in"
      Columns(8).NumberFormat=   "hh:nn:ss"
      Columns(8).ExternalEditor=   "txtTime"
      Columns(8).ExternalEditor.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "2nd Time Out"
      Columns(9).DataField=   "t2out"
      Columns(9).NumberFormat=   "hh:nn:ss"
      Columns(9).ExternalEditor=   "txtTime"
      Columns(9).ExternalEditor.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Shift Code"
      Columns(10).DataField=   "shiftcode"
      Columns(10).DropDown=   "tddShift"
      Columns(10).DropDown.vt=   8
      Columns(10).ExternalEditor=   "txtShiftcode"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Shift Detail"
      Columns(11).DataField=   "shiftdetail"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Hours Worked"
      Columns(12).DataField=   "wrkhrs"
      Columns(12).NumberFormat=   "#,##0.00"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Absent"
      Columns(13).DataField=   "absent"
      Columns(13).NumberFormat=   "#,##0.00"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Late Hours"
      Columns(14).DataField=   "latehrs"
      Columns(14).NumberFormat=   "#,##0.00"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Undertime Hours"
      Columns(15).DataField=   "uthrs"
      Columns(15).NumberFormat=   "#,##0.00"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1561"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1482"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=926"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=847"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=979"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=900"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=953"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=873"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1296"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1217"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=1323"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=1244"
      Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=1296"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1217"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=1296"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=1217"
      Splits(0)._ColumnProps(55)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(10).Width=1482"
      Splits(0)._ColumnProps(58)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(10)._WidthInPix=1402"
      Splits(0)._ColumnProps(60)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(61)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(62)=   "Column(11).Width=4154"
      Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=4075"
      Splits(0)._ColumnProps(65)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(67)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(68)=   "Column(12).Width=1217"
      Splits(0)._ColumnProps(69)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(12)._WidthInPix=1138"
      Splits(0)._ColumnProps(71)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(73)=   "Column(13).Width=1217"
      Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=1138"
      Splits(0)._ColumnProps(76)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(77)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(78)=   "Column(14).Width=1191"
      Splits(0)._ColumnProps(79)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(14)._WidthInPix=1111"
      Splits(0)._ColumnProps(81)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(82)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(83)=   "Column(15).Width=1429"
      Splits(0)._ColumnProps(84)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(15)._WidthInPix=1349"
      Splits(0)._ColumnProps(86)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(87)=   "Column(15).Order=16"
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
      HeadLines       =   2
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=74,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=86,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=102,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=94,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=91,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=92,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=93,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=87,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=88,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=89,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=62,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=78,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=75,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=76,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=77,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=82,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
      _StyleDefs(98)  =   "Named:id=33:Normal"
      _StyleDefs(99)  =   ":id=33,.parent=0"
      _StyleDefs(100) =   "Named:id=34:Heading"
      _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   ":id=34,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=35:Footing"
      _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(105) =   "Named:id=36:Selected"
      _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=37:Caption"
      _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(109) =   "Named:id=38:HighlightRow"
      _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(111) =   "Named:id=39:EvenRow"
      _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(113) =   "Named:id=40:OddRow"
      _StyleDefs(114) =   ":id=40,.parent=33"
      _StyleDefs(115) =   "Named:id=41:RecordSelector"
      _StyleDefs(116) =   ":id=41,.parent=34"
      _StyleDefs(117) =   "Named:id=42:FilterBar"
      _StyleDefs(118) =   ":id=42,.parent=33"
   End
   Begin CitronSoftwarePayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   45
      TabIndex        =   9
      Top             =   90
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Generate DTR"
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
      ForeColor       =   3186872
      GradTheme       =   2
   End
End
Attribute VB_Name = "frmDtr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTitoTmp       As ADODB.Recordset
Dim rsDtrTmp        As ADODB.Recordset

Private Sub cmdGenerate_Click()

    Dim mTin        As Date
    Dim mTout       As Date
    Dim mLasTout    As String
    Dim mAdvDate    As String
    
    Dim mT1in       As String
    Dim mT1out      As String
    Dim mT2in       As String
    Dim mT2out      As String
    Dim mST1in      As String
    Dim mST1out     As String
    Dim mST2in      As String
    Dim mST2out     As String
    Dim mWrkdate    As String
    
    Dim mWrkHrs     As Double
    
    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        MsgBox "Please select a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbEmployee.Text) = "" Or IsNull(tdbEmployee.SelectedItem) Or tdbEmployee.ApproxCount = 0 Then
        MsgBox "Please select an employee.", vbExclamation + vbOKOnly
        tdbEmployee.SetFocus
        Exit Sub
    End If
                    
    Create_TmpTito
    
    Load_Dtr
    
    
    If rsDtrTmp.RecordCount > 0 Then
        'Assigning TITO to Employee DTR
        
            With rsDtrTmp
                .MoveFirst
                Do While Not .EOF
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
                                    'compute for late
                                    If CDate(mT1in) > CDate(mST1in) Then
                                        If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                            .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in))
                                        End If
                                    End If
                                    'check if late on 2nd time in
                                    If !brkhrsperday > 0 Then
                                        If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                            If .Fields("latehrs") > 0 Then
                                                .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday)
                                            Else
                                                .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday
                                            End If
                                        End If
                                    End If
                                    'compute for undertime
                                    If CDate(mT2out) < CDate(mST2out) Then
                                        If DiffHrs(CDate(mT2out), CDate(mST2out)) > 0 Then
                                            .Fields("uthrs") = DiffHrs(CDate(mT2out), CDate(mST2out))
                                        End If
                                    End If
                                    mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                ElseIf mT1in <> "" And mT2in = "" Then 'if only the first two (2) time slots were consumed.
                                    If CDate(mT1in) <= CDate(mST1in) Then
                                        mTin = mST1in
                                    Else
                                        mTin = mT1in
                                        'compute for late
                                        If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                            .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in))
                                        End If
                                    End If
                                    If CDate(mT1out) < CDate(mST1out) Then
                                        mTout = mT1out
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                        'compute for 1st undertime
                                        If DiffHrs(CDate(mT1out), CDate(mST1out)) > 0 Then
                                            .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                        End If
                                        If Trim(.Fields("holiday")) = "" Then .Fields("absent") = 0.5
                                    ElseIf CDate(mT1out) >= CDate(mST1out) And CDate(mT1out) < CDate(mST2in) Then
                                        mTout = mST1out
                                        mWrkHrs = DiffHrs(mTin, mTout)
                                        If Trim(.Fields("holiday")) = "" Then .Fields("absent") = 0.5
                                    ElseIf CDate(mT1out) >= CDate(mST2in) And CDate(mT1out) < CDate(mST2out) Then
                                        mTout = mT1out
                                        'compute for 2nd undertime
                                        If DiffHrs(CDate(mT1out), CDate(mST2out)) > 0 Then
                                            .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST2out))
                                        End If
                                        mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    Else
                                        mTout = mST2out
                                        mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                    End If
                                ElseIf mT1in = "" And mT2in <> "" Then ' if only the last two(2) time slots were consumed.
                                    If CDate(mT2in) <= CDate(mST2in) Then
                                        mTin = mST2in
                                    Else
                                        mTin = mT2in
                                        If DiffHrs(CDate(mST1in), CDate(mT2in)) > 0 Then
                                            .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in))
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
                                    If Trim(.Fields("holiday")) = "" Then .Fields("absent") = 0.5
                                    mWrkHrs = DiffHrs(mTin, mTout)
                                Else
                                    If Trim(.Fields("holiday")) = "" Then .Fields("absent") = 1
                                End If
                            Else 'On travel or On leave
                                'If on travel or on leave during the second shift, compute only the first shift.
                                If !firsttravel = 0 And !firstleave = 0 Then
                                    If mT1in <> "" And mT2in <> "" Then
                                        'compute for late
                                        If CDate(mT1in) > CDate(mST1in) Then
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in))
                                            End If
                                        End If
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            If CDate(mT2in) < CDate(mST1out) Then
                                                'compute for late
                                                If !brkhrsperday > 0 Then
                                                    If DiffHrs(CDate(mT1out), CDate(mT2in)) > !brkhrsperday Then
                                                        If .Fields("latehrs") > 0 Then
                                                            .Fields("latehrs") = .Fields("latehrs") + (DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday)
                                                        Else
                                                            .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in)) - !brkhrsperday
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
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                        End If
                                    ElseIf mT1in <> "" And mT2in = "" Then
                                        'compute for late
                                        If CDate(mT1in) > CDate(mST1in) Then
                                            If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                                .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in))
                                            End If
                                        End If
                                        'Compute for undertime
                                        If CDate(mT1out) < CDate(mST1out) Then
                                            .Fields("uthrs") = DiffHrs(CDate(mT1out), CDate(mST1out))
                                        End If
                                        mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out)) - .Fields("uthrs") - .Fields("latehrs")
                                        If !secondtravel = 1 Then
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                        End If
                                    ElseIf mT1in = "" And mT2in <> "" Then
                                        .Fields("absent") = 0.5
                                        If !secondtravel = 1 Then
                                            mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                        End If
                                    End If
                                'If on travel or on leave during the first shift, compute only the second shift
                                ElseIf !secondtravel = 0 And !secondleave = 0 Then
                                    If mT1in <> "" And mT2in <> "" Then
                                        'compute for late
                                        If CDate(mT1out) >= CDate(mST2in) Then
                                            If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                .Fields("latehrs") = DiffHrs(CDate(mT1out), CDate(mT2in))
                                            End If
                                        Else
                                            If DiffHrs(CDate(mT1out), CDate(mT2in)) > 0 Then
                                                If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                    .Fields("latehrs") = DiffHrs(CDate(mT2in), CDate(mST1out))
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
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
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
                                                    mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                End If
                                            ElseIf CDate(mT1out) > CDate(mST2out) Then
                                                mWrkHrs = DiffHrs(CDate(mST2in), CDate(mST2out))
                                                If !firsttravel = 1 Then
                                                    mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                                End If
                                            ElseIf CDate(mT1out) <= CDate(mST2in) Then
                                                .Fields("absent") = 0.5
                                                If !firsttravel = 1 Then
                                                    mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                                End If
                                            End If
                                        End If
                                    ElseIf mT1in = "" And mT2in <> "" Then
                                        'compute for late
                                        If CDate(mT2in) > CDate(mST2in) Then
                                            If DiffHrs(CDate(mST2in), CDate(mT2in)) > 0 Then
                                                .Fields("latehrs") = DiffHrs(CDate(mST2in), CDate(mT2in))
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
                                            mWrkHrs = mWrkHrs + DiffHrs(CDate(mST1in), CDate(mST1out))
                                        End If
                                    End If
                                Else
                                    If !firsttravel = 1 Then
                                        mWrkHrs = DiffHrs(CDate(mST1in), CDate(mST1out))
                                    End If
                                    If !secondtravel = 1 Then
                                        mWrkHrs = mWrkHrs + DiffHrs(CDate(mST2in), CDate(mST2out))
                                    End If
                                End If
                            End If
                        Else 'For schedules with only two(2) time slots
                            If !travel = 0 And !leave = 0 Then
                                If mT1in <> "" Then 'Check if time slots were used.
                                    If CDate(mT1in) < CDate(mST1in) Then
                                        mTin = mST1in
                                    Else
                                        mTin = mT1in
                                        'compute for late
                                        If DiffHrs(CDate(mST1in), CDate(mT1in)) > 0 Then
                                            .Fields("latehrs") = DiffHrs(CDate(mST1in), CDate(mT1in))
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
                                    mWrkHrs = DiffHrs(mTin, mTout)
                                    'mWrkHrs = !hrsperday - .Fields("uthrs") - .Fields("latehrs")
                                Else 'absent
                                    If Trim(.Fields("holiday")) = "" Then .Fields("absent") = 1
                                End If
                            End If
                        End If
                        .Fields("wrkhrs") = mWrkHrs
                    End If
                    
                    
                    
                    .MoveNext
                    DoEvents
                Loop
                .MoveFirst
                
            End With
        
        
    Else

    End If
    
End Sub

Private Sub Form_Load()
  
    Add_MDIButton Me.Name, TitleBar.Caption
  
    bind_tdb CitronPayroll, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto from payrollperiod", "description", "percode"
    
    bind_tdb CitronPayroll, tdbEmployee, "select empno,concat(lastname, ', ', firstname,' ',middlename) fullname from employee order by concat(lastname, ', ', firstname,' ',middlename)", "fullname", "empno"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fra1
        .Top = TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With tdgTito
        .Top = fra1.Height + TitleBar.Height
        .Left = 0
        .Height = Me.ScaleHeight - (fra1.Height + TitleBar.Height)
    End With
    
    With tdgDtr
        .Top = fra1.Height + TitleBar.Height
        .Left = tdgTito.Width
        .Height = Me.ScaleHeight - (fra1.Height + TitleBar.Height)
        .Width = Me.ScaleWidth - tdgTito.Width
    End With

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

    Set tdgTito.DataSource = rsTitoTmp

    NetOpen rsTito, "", "select empno,complog,datelog,timelog,logstat " & _
                        "from tito where empno = '" & tdbEmployee.BoundText & "' and " & _
                        "datelog Between '" & Format(tdbPayrollPeriod.Columns("from").Text - 1, "YYYY-MM-DD") & "'  and " & _
                        "'" & Format(tdbPayrollPeriod.Columns("to").Text + 1, "YYYY-MM-DD") & "' " & _
                        "Union All " & _
                        "select empno,complog,datelog,timelog,logstat " & _
                        "from gplne where empno = '" & tdbEmployee.BoundText & "' and " & _
                        "(datelog Between '" & Format(tdbPayrollPeriod.Columns("from").Text - 1, "YYYY-MM-DD") & "' and " & _
                        "'" & Format(tdbPayrollPeriod.Columns("to").Text + 1, "YYYY-MM-DD") & "') and  " & _
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

        End If

    End With
    
End Sub

Private Sub Load_Dtr()

    Dim mdate       As Date
    Dim rsEmpDtr    As ADODB.Recordset
    Dim rsEmpShift  As ADODB.Recordset
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
        .Fields.Append "absent", adDouble
        .Fields.Append "latehrs", adDouble
        .Fields.Append "uthrs", adDouble
        .Fields.Append "hrsperday", adDouble, 18
        .Fields.Append "brkhrsperday", adDouble, 18
        .Fields.Append "firsttravel", adVarChar, 1
        .Fields.Append "secondtravel", adVarChar, 1
        .Fields.Append "firstleave", adVarChar, 1
        .Fields.Append "secondleave", adVarChar, 1
        .Open
        
        Set tdgDtr.DataSource = rsDtrTmp

        mdate = Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "MM/DD/YYYY")
        
        Do While mdate <= Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "MM/DD/YYYY")
          
            .AddNew
            .Fields("wrkdate") = mdate
            .Fields("dayno") = Weekday(mdate)
            .Fields("day") = WeekdayName(Weekday(mdate))
            
            NetOpen rsEmpDtr, "", "select * from dtremp where empno = '" & tdbEmployee.BoundText & "' and " & _
                                "workdate = '" & Format(mdate, "YYYY-MM-DD") & "'"
                                
            If rsEmpDtr.RecordCount > 0 Then
              .Fields("updatable") = IIf(rsEmpDtr!updatable = "Y", 1, 0)
              .Fields("t1in") = rsEmpDtr!t1in
              .Fields("t1out") = rsEmpDtr!t1out
              .Fields("t2in") = rsEmpDtr!t2in
              .Fields("t2out") = rsEmpDtr!t2out
              .Fields("st1in") = rsEmpDtr!st1in
              .Fields("st1out") = rsEmpDtr!st1out
              .Fields("st2in") = rsEmpDtr!st2in
              .Fields("st2out") = rsEmpDtr!st2out
              .Fields("shiftcode") = rsEmpDtr!shiftcode
              .Fields("shiftdetail") = rsEmpDtr!st1in & "   " & rsEmpDtr!st1out & "       " & rsEmpDtr!st2in & "   " & rsEmpDtr!st2out
              .Fields("brkstart") = rsEmpDtr!brkstart
              .Fields("brkend") = rsEmpDtr!brkend
              .Fields("nitepremstart") = rsEmpDtr!nitepremstart
              .Fields("nitepremend") = rsEmpDtr!nitepremend
              .Fields("hrsperday") = rsEmpDtr!hrsperday
              .Fields("brkhrsperday") = rsEmpDtr!brkhrsperday
              .Fields("dayoff") = IIf(rsEmpDtr!dayoff = "Y", 1, 0)
              
            Else
              
              NetOpen rsEmpShift, "", "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, brkhrs brkhrsperday from empshift x1 left outer join shift x2 on " & _
                                    "x1.shiftcode = x2.shiftcode where x1.shiftcode <> '' and x1.empno = '" & tdbEmployee.BoundText & "' and x1.dayno = '" & Weekday(mdate) & "'"
              
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
                    
              Else
                    .Fields("updatable") = 1
                    .Fields("dayoff") = 1
              End If
              
            End If
            
            NetOpen rsHoliday, "", "select * from holiday where holidaydate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
            If rsHoliday.RecordCount > 0 Then
                If CInt(rsHoliday!regular) = 1 Then
                  .Fields("holiday") = "Legal"
                Else
                  .Fields("holiday") = "Special"
                End If
            Else
                .Fields("holiday") = ""
            End If
            
            NetOpen rsOBT, "", "select x1.* from obtlne x1 left outer join obthdr x2 on x1.obtnum = x2.obtnum " & _
                             "where x2.empno = '" & tdbEmployee.BoundText & "' and x1.obtdate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
            If rsOBT.RecordCount > 0 Then
                .Fields("travel") = 1
                .Fields("firsttravel") = rsOBT!firstshift
                .Fields("secondtravel") = rsOBT!secondshift
            Else
                .Fields("travel") = 0
                .Fields("firsttravel") = 0
                .Fields("secondtravel") = 0
            End If
            
            NetOpen rsLeave, "", "select x1.* from lvlne x1 " & _
                    "left outer join lvhdr x2 on x1.lvnum = x2.lvnum " & _
                    "where x2.empno = '" & tdbEmployee.BoundText & "' and x1.lvdate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
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
            
            mdate = mdate + 1
            DoEvents
        Loop
      
    End With

End Sub

Private Sub cmdImport_Click()
    
    Dim rsFBTito    As ADODB.Recordset
    Dim mTime       As String
    
    Cnstr.CommandTimeout = 2000
    Set rsFBTito = New ADODB.Recordset
    rsFBTito.Open "select * from tito where empno = 'FBA26506'", Cnstr, adOpenKeyset, adLockOptimistic
    
    With rsFBTito
        If .RecordCount > 0 Then
            .MoveFirst
            pb.Max = .RecordCount
            pb.Value = 0
            Do While Not .EOF
                pb.Value = pb.Value + 1
                
                CitronPayroll.Execute "insert into tito(empno,biometid,complog,datelog,timelog,logstat,cancelled,remarks) values " & _
                        "('0000002','000002','" & Format(CDate(!titodate) + 365 & " " & CDate(!titotime), "YYYY-MM-DD HH:NN:SS") & "', " & _
                        "'" & Format(CDate(!titodate) + 365, "YYYY-MM-DD") & "','" & Format(!titotime, "hh:nn:ss") & "', " & _
                        "'" & IIf(!Type = "I", "In", "Out") & "','N','')"

'                mTime = Format(CDate(!titodate & " " & !titotime) - CDate("12:00:00"), "HH:NN:SS")
'
'                CitronPayroll.Execute "insert into tito(empno,biometid,complog,datelog,timelog,logstat,cancelled,remarks) values " & _
'                        "('0000004','000004','" & Format(CDate(!titodate) & " " & mTime, "YYYY-MM-DD HH:NN:SS") & "', " & _
'                        "'" & Format(CDate(!titodate), "YYYY-MM-DD") & "','" & Format(mTime, "hh:nn:ss") & "', " & _
'                        "'" & IIf(!Type = "I", "In", "Out") & "','N','')"

                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
End Sub

Private Function DiffHrs(mHrs1 As Date, mHrs2 As Date) As Double
    DiffHrs = Format(Round(DateDiff("N", mHrs1, mHrs2) / 60, 2), "#,##0.00")
End Function


