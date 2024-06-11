VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmComputeOvertime 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   13350
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   105
      TabIndex        =   4
      Top             =   1005
      Width           =   12720
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   345
         Left            =   1515
         TabIndex        =   5
         Tag             =   "Municipal"
         Top             =   465
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
         _PropDict       =   $"frmComputeOvertime.frx":0000
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
         Left            =   1515
         TabIndex        =   6
         Tag             =   "Municipal"
         Top             =   105
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
         _PropDict       =   $"frmComputeOvertime.frx":00AA
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
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   525
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   135
         Width           =   1830
      End
   End
   Begin CitronSoftwarePayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   60
      TabIndex        =   0
      Top             =   375
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      BorderColor     =   14215660
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   420
         Left            =   75
         TabIndex        =   1
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Generate"
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
   End
   Begin CitronSoftwarePayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Employee's Overtime"
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
   Begin TrueOleDBGrid80.TDBGrid tdgOTtito 
      Height          =   3060
      Left            =   0
      TabIndex        =   3
      Top             =   2145
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
      Appearance      =   3
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
   Begin TrueOleDBGrid80.TDBGrid tdgOTDtr 
      Height          =   3210
      Left            =   5460
      TabIndex        =   9
      Top             =   2130
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
      Columns(3).Caption=   "Day Off"
      Columns(3).DataField=   "dayoff"
      Columns(3).DataWidth=   1
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Actual IN"
      Columns(4).DataField=   "actotstart"
      Columns(4).NumberFormat=   "hh:nn:ss"
      Columns(4).ExternalEditor=   "txtTime"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Actual OUT"
      Columns(5).DataField=   "actotend"
      Columns(5).NumberFormat=   "hh:nn:ss"
      Columns(5).ExternalEditor=   "txtTime"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Overtime IN"
      Columns(6).DataField=   "otstart"
      Columns(6).NumberFormat=   "hh:nn:ss"
      Columns(6).ExternalEditor=   "txtTime"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Overtitme OUT"
      Columns(7).DataField=   "otend"
      Columns(7).NumberFormat=   "hh:nn:ss"
      Columns(7).ExternalEditor=   "txtTime"
      Columns(7).ExternalEditor.vt=   8
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Hours Worked"
      Columns(8).DataField=   "otwrkhrs"
      Columns(8).NumberFormat=   "#,##0.00"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Night Premium Hours Worked"
      Columns(9).DataField=   "NiteWrkHrs"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
      Splits(0)._ColumnProps(19)=   "Column(3).Width=953"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=873"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1720"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1640"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=1746"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=1799"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=1720"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=2037"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1958"
      Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(45)=   "Column(8).Width=1217"
      Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(48)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(53)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   3
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
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
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
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=102,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=99,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=100,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=17"
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
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmComputeOvertime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsOTtmp         As ADODB.Recordset
Dim rsTitoTmp       As ADODB.Recordset

Private Sub Form_Load()

    bind_tdb CitronPayroll, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto from payrollperiod", "description", "percode"
    
    bind_tdb CitronPayroll, tdbEmployee, "select empno,concat(lastname, ', ', firstname,' ',middlename) fullname from employee order by concat(lastname, ', ', firstname,' ',middlename)", "fullname", "empno"
      
    Add_MDIButton Me.Name, TitleBar.Caption
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fra1
        .Top = frabutton.Height + frabutton.Top
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With tdgOTtito
        .Top = fra1.Height + fra1.Top
        .Left = 0
        .Height = Me.ScaleHeight - (fra1.Height + TitleBar.Height)
    End With
    
    With tdgOTDtr
        .Top = fra1.Height + fra1.Top
        .Left = tdgOTtito.Width
        .Height = Me.ScaleHeight - (fra1.Height + TitleBar.Height)
        .Width = Me.ScaleWidth - tdgOTtito.Width
    End With
    
End Sub

Private Sub cmdGenerate_Click()
    
    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        MsgBox "Please choose a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbEmployee.Text) = "" Or IsNull(tdbEmployee.SelectedItem) Or tdbEmployee.ApproxCount = 0 Then
        MsgBox "Please choose an emloyee.", vbExclamation + vbOKOnly
        tdbEmployee.SetFocus
        Exit Sub
    End If
    
    Create_OTDtr
    
End Sub

Private Sub Create_OTDtr()
    
    Dim rsOT        As ADODB.Recordset
    Dim rsDtrEmp    As ADODB.Recordset
    Dim rsHoliday   As ADODB.Recordset
    
    'Time Reference
    Dim mOT_Start   As String
    Dim mOT_End     As String
    Dim mNiteStart  As String
    Dim mNiteEnd    As String
    Dim mTimeIN     As String
    Dim mTimeOUT    As String
    
    'Actual Time
    Dim mOTActIN    As Date
    Dim mOTActOUT   As Date
    Dim mNiteIN     As Date
    Dim mNiteOUT    As Date
    
    Set rsOTtmp = Nothing
    Set rsOTtmp = New ADODB.Recordset
    
    With rsOTtmp
        .Fields.Append "otcode", adVarChar, 15
        .Fields.Append "empno", adVarChar, 15
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
        .Fields.Append "nitewrkhrs", adDouble, 18
        .Open
    End With
    
    Set tdgOTDtr.DataSource = rsOTtmp
    
    NetOpen rsOT, "", "select * from overtime where empno = '" & tdbEmployee.BoundText & "' and status = 'Approved' order by wrkdate,otstart"
    
    With rsOTtmp
        If rsOT.RecordCount > 0 Then
            rsOT.MoveFirst
            Do While Not rsOT.EOF
            
            
            mOT_Start = ""
            mOT_End = ""
            mNiteStart = ""
            mNiteEnd = ""
            mTimeIN = ""
            mTimeOUT = ""
            
            .AddNew
            .Fields("otcode") = rsOT!otcode
            .Fields("empno") = rsOT!empno
            .Fields("percode") = rsOT!percode
            .Fields("wrkdate") = rsOT!wrkdate
            .Fields("day") = WeekdayName(Weekday(rsOT!wrkdate))
            .Fields("otstart") = rsOT!otstart
            .Fields("otend") = rsOT!otend
            
            'Check if dayoff
            NetOpen rsDtrEmp, "", "select * from dtremp where workdate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
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
            
            'Check if Holiday
            NetOpen rsHoliday, "", "select * from holiday where holidaydate = '" & Format(rsOT!wrkdate, "YYYY-MM-DD") & "'"
            If rsHoliday.RecordCount > 0 Then
                If CInt(rsHoliday!regular) = 1 Then
                  .Fields("holiday") = "Legal"
                Else
                  .Fields("holiday") = "Special"
                End If
            Else
                .Fields("holiday") = ""
            End If
            
            Create_TmpOTTito
            
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
                !nitewrkhrs = 0
                
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
            
            
            
            .Update
            rsOT.MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub Create_TmpOTTito()
    
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

    Set tdgOTtito.DataSource = rsTitoTmp

    NetOpen rsTito, "", "select * from tito where empno = '" & tdbEmployee.BoundText & "' and datelog between " & _
                        "'" & Format(rsOTtmp!wrkdate - 1, "YYYY-MM-DD") & "' and " & _
                        "'" & Format(rsOTtmp!wrkdate + 1, "YYYY-MM-DD") & "' order by complog"

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

Private Function DiffHrs(mHrs1 As Date, mHrs2 As Date) As Double
    DiffHrs = Format(Round(DateDiff("N", mHrs1, mHrs2) / 60, 2), "#,##0.00")
End Function


