VERSION 5.00
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDtr 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   13095
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      Height          =   705
      Left            =   75
      TabIndex        =   3
      Top             =   1050
      Width           =   12720
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   315
         Left            =   6885
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   255
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   556
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
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
         EditFont        =   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _PropDict       =   $"frmDtr.frx":0000
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
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
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   255
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   556
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1879"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2355"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2249"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2752"
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
         EditFont        =   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _PropDict       =   $"frmDtr.frx":00AA
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
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
         Height          =   225
         Left            =   5970
         TabIndex        =   5
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Width           =   1830
      End
   End
   Begin TDBTime6Ctl.TDBTime txtTime 
      Height          =   285
      Left            =   4785
      TabIndex        =   6
      Top             =   7515
      Visible         =   0   'False
      Width           =   1710
      _Version        =   65536
      _ExtentX        =   3016
      _ExtentY        =   503
      Caption         =   "frmDtr.frx":0154
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmDtr.frx":01C0
      Spin            =   "frmDtr.frx":0210
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn AM/PM"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn AMPM"
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
      Text            =   "11:01 AM"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0.459039351851852
   End
   Begin TrueOleDBGrid80.TDBDropDown tddShift 
      Height          =   1365
      Left            =   165
      TabIndex        =   7
      Top             =   6435
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Shift Code"
      Columns(0).DataField=   "shiftcode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Shift Description"
      Columns(1).DataField=   "shiftdesc"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "t1in"
      Columns(2).DataField=   "t1in"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "t1out"
      Columns(3).DataField=   "t1out"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "t2in"
      Columns(4).DataField=   "t2in"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "t2out"
      Columns(5).DataField=   "t2out"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "brkstart"
      Columns(6).DataField=   "brkstart"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "brkend"
      Columns(7).DataField=   "brkend"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "nitepremstart"
      Columns(8).DataField=   "nitepremstart"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "nitepremend"
      Columns(9).DataField=   "nitepremend"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "hrsperday"
      Columns(10).DataField=   "hrsperday"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "brkhrsperday"
      Columns(11).DataField=   "brkhrsperday"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(57)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(59)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(63)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(65)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(69)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(11).Order=12"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   "Branch"
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   -1  'True
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   16185592
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
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HDAFAEF&"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid tdgDtr 
      Height          =   5880
      Left            =   8835
      TabIndex        =   2
      Top             =   3360
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   10372
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Updatable"
      Columns(0).DataField=   "updatable"
      Columns(0).NumberFormat=   "True/False"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Date"
      Columns(1).DataField=   "wrkdate"
      Columns(1).NumberFormat=   "MM/DD/YYYY"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Day #"
      Columns(2).DataField=   "dayno"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Day"
      Columns(3).DataField=   "day"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Holiday"
      Columns(4).DataField=   "holiday"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "On Travel"
      Columns(5).DataField=   "travel"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "On Leave"
      Columns(6).DataField=   "leave"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   4
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Day Off"
      Columns(7).DataField=   "dayoff"
      Columns(7).DataWidth=   1
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "1st Time In"
      Columns(8).DataField=   "t1in"
      Columns(8).NumberFormat=   "hh:nn AM/PM"
      Columns(8).ExternalEditor=   "txtTime"
      Columns(8).ExternalEditor.vt=   8
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "1st Time Out"
      Columns(9).DataField=   "t1Out"
      Columns(9).NumberFormat=   "hh:nn AM/PM"
      Columns(9).ExternalEditor=   "txtTime"
      Columns(9).ExternalEditor.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "2nd Time In"
      Columns(10).DataField=   "t2in"
      Columns(10).NumberFormat=   "hh:nn AM/PM"
      Columns(10).ExternalEditor=   "txtTime"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "2nd Time Out"
      Columns(11).DataField=   "t2out"
      Columns(11).NumberFormat=   "hh:nn AM/PM"
      Columns(11).ExternalEditor=   "txtTime"
      Columns(11).ExternalEditor.vt=   8
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Shift Code"
      Columns(12).DataField=   "shiftcode"
      Columns(12).DropDown=   "tddShift"
      Columns(12).DropDown.vt=   8
      Columns(12).ExternalEditor=   "txtShiftcode"
      Columns(12).ExternalEditor.vt=   8
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Shift Detail"
      Columns(13).DataField=   "shiftdetail"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "st1in"
      Columns(14).DataField=   "st1in"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "st1out"
      Columns(15).DataField=   "st1out"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "st2in"
      Columns(16).DataField=   "st2in"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "st2out"
      Columns(17).DataField=   "st2out"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "brkstart"
      Columns(18).DataField=   "brkstart"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "brkend"
      Columns(19).DataField=   "brkend"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "nitepremstart"
      Columns(20).DataField=   "nitepremstart"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "nitepremend"
      Columns(21).DataField=   "nitepremend"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "hrsperday"
      Columns(22).DataField=   "hrsperday"
      Columns(22).NumberFormat=   "#,##0.00"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "brkhrsperday"
      Columns(23).DataField=   "brkhrsperday"
      Columns(23).NumberFormat=   "#,##0.00"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   24
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=24"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1482"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1402"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1693"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=900"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=820"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(18)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1984"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1905"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1561"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1482"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=926"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=847"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=8193"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=953"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=873"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=873"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=794"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=1826"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1746"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=1852"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1773"
      Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(60)=   "Column(10).Width=1905"
      Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=1826"
      Splits(0)._ColumnProps(63)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(64)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(65)=   "Column(11).Width=2037"
      Splits(0)._ColumnProps(66)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(11)._WidthInPix=1958"
      Splits(0)._ColumnProps(68)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(69)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(70)=   "Column(12).Width=1482"
      Splits(0)._ColumnProps(71)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(12)._WidthInPix=1402"
      Splits(0)._ColumnProps(73)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(74)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(75)=   "Column(13).Width=5106"
      Splits(0)._ColumnProps(76)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(13)._WidthInPix=5027"
      Splits(0)._ColumnProps(78)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(79)=   "Column(13)._ColStyle=8196"
      Splits(0)._ColumnProps(80)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(81)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(82)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(84)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(85)=   "Column(14)._ColStyle=8196"
      Splits(0)._ColumnProps(86)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(87)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(88)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(89)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(91)=   "Column(15)._EditAlways=0"
      Splits(0)._ColumnProps(92)=   "Column(15)._ColStyle=8196"
      Splits(0)._ColumnProps(93)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(95)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(96)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(98)=   "Column(16)._EditAlways=0"
      Splits(0)._ColumnProps(99)=   "Column(16)._ColStyle=8196"
      Splits(0)._ColumnProps(100)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(101)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(102)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(103)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(104)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(105)=   "Column(17)._EditAlways=0"
      Splits(0)._ColumnProps(106)=   "Column(17)._ColStyle=8196"
      Splits(0)._ColumnProps(107)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(108)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(109)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(18)._EditAlways=0"
      Splits(0)._ColumnProps(113)=   "Column(18)._ColStyle=8196"
      Splits(0)._ColumnProps(114)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(115)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(116)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(117)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(118)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(119)=   "Column(19)._EditAlways=0"
      Splits(0)._ColumnProps(120)=   "Column(19)._ColStyle=8196"
      Splits(0)._ColumnProps(121)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(122)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(123)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(124)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(125)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(126)=   "Column(20)._EditAlways=0"
      Splits(0)._ColumnProps(127)=   "Column(20)._ColStyle=8196"
      Splits(0)._ColumnProps(128)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(129)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(130)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(131)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(132)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(133)=   "Column(21)._EditAlways=0"
      Splits(0)._ColumnProps(134)=   "Column(21)._ColStyle=8196"
      Splits(0)._ColumnProps(135)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(136)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(137)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(138)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(140)=   "Column(22)._EditAlways=0"
      Splits(0)._ColumnProps(141)=   "Column(22)._ColStyle=8196"
      Splits(0)._ColumnProps(142)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(143)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(144)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(145)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(146)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(147)=   "Column(23)._EditAlways=0"
      Splits(0)._ColumnProps(148)=   "Column(23)._ColStyle=8196"
      Splits(0)._ColumnProps(149)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(150)=   "Column(23).Order=24"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=98,.parent=13,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=74,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=102,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=46,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=66,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=63,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=64,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=65,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=94,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=91,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=92,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=93,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=106,.parent=13,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=103,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=104,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=105,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=110,.parent=13,.locked=-1"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=107,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=108,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=109,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=114,.parent=13,.locked=-1"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=111,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=112,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=113,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=118,.parent=13,.locked=-1"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=115,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=116,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=117,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=75,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=76,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=77,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=82,.parent=13,.locked=-1"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=79,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=80,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=81,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=86,.parent=13,.locked=-1"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=83,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=84,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=85,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=122,.parent=13,.locked=-1"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=126,.parent=13,.locked=-1"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(23).Style:id=130,.parent=13,.locked=-1"
      _StyleDefs(127) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
      _StyleDefs(130) =   "Named:id=33:Normal"
      _StyleDefs(131) =   ":id=33,.parent=0"
      _StyleDefs(132) =   "Named:id=34:Heading"
      _StyleDefs(133) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(134) =   ":id=34,.wraptext=-1"
      _StyleDefs(135) =   "Named:id=35:Footing"
      _StyleDefs(136) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(137) =   "Named:id=36:Selected"
      _StyleDefs(138) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(139) =   "Named:id=37:Caption"
      _StyleDefs(140) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(141) =   "Named:id=38:HighlightRow"
      _StyleDefs(142) =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(143) =   "Named:id=39:EvenRow"
      _StyleDefs(144) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(145) =   "Named:id=40:OddRow"
      _StyleDefs(146) =   ":id=40,.parent=33"
      _StyleDefs(147) =   "Named:id=41:RecordSelector"
      _StyleDefs(148) =   ":id=41,.parent=34"
      _StyleDefs(149) =   "Named:id=42:FilterBar"
      _StyleDefs(150) =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText txtShiftcode 
      Height          =   285
      Left            =   6510
      TabIndex        =   8
      Top             =   7515
      Visible         =   0   'False
      Width           =   2130
      _Version        =   65536
      _ExtentX        =   3757
      _ExtentY        =   503
      Caption         =   "frmDtr.frx":0238
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDtr.frx":02A4
      Key             =   "frmDtr.frx":02C2
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
   Begin TDBText6Ctl.TDBText txtDayStat 
      Height          =   255
      Left            =   9330
      TabIndex        =   9
      Top             =   7485
      Visible         =   0   'False
      Width           =   2130
      _Version        =   65536
      _ExtentX        =   3757
      _ExtentY        =   450
      Caption         =   "frmDtr.frx":0306
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDtr.frx":0372
      Key             =   "frmDtr.frx":0390
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
      Appearance      =   2
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
      MaxLength       =   1
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
   Begin CitronSoftwarePayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   120
      TabIndex        =   10
      Top             =   435
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      BorderColor     =   14215660
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   420
         Left            =   1230
         TabIndex        =   11
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Save"
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
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   420
         Left            =   75
         TabIndex        =   12
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
      TabIndex        =   13
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Generate Emloyee's DTR"
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
Attribute VB_Name = "frmDtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDtr           As ADODB.Recordset
Dim rsDtrTmp        As ADODB.Recordset
Dim mPercode        As String
Dim mEmpNo          As String

Private Sub cmdGenerate_Click()
  
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
    
    Load_Dtr
    
    If rsDtrTmp.RecordCount > 0 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If

End Sub

Private Sub cmdSave_Click()
    
    Dim mAdvCnt     As Integer

    With rsDtrTmp
    
        If .RecordCount > 0 Then
        
            .MoveFirst
            
            Do While Not .EOF
            
                mAdvCnt = 0
                If IsDate(!t1in) And IsDate(!t1out) And IsDate(!t2in) And IsDate(!t2out) Then
                    If CDate(!t1in) > CDate(!t1out) Then
                        mAdvCnt = 1
                    End If
                    If CDate(!t1out) > CDate(!t2in) Then
                        mAdvCnt = mAdvCnt + 1
                        If mAdvCnt > 1 Then
                            MsgBox "Invalid time.", vbExclamation + vbOKOnly
                            tdgDtr.SetFocus
                            tdgDtr.Col = tdgDtr.Columns("t2in").ColIndex
                            Exit Sub
                        End If
                    End If
                    If CDate(!t2in) > CDate(!t2out) Then
                        mAdvCnt = mAdvCnt + 1
                        If mAdvCnt > 1 Then
                            MsgBox "Invalid time.", vbExclamation + vbOKOnly
                            tdgDtr.SetFocus
                            tdgDtr.Col = tdgDtr.Columns("t2out").ColIndex
                            Exit Sub
                        End If
                    End If
                End If
                
                If !dayoff = 0 Then
                    If Trim(!shiftcode) = "" Then
                        MsgBox "Please select a shift schedule.", vbExclamation + vbOKOnly
                        tdgDtr.SetFocus
                        tdgDtr.Col = tdgDtr.Columns("shiftcode").ColIndex
                        Exit Sub
                    End If
                End If
                
                .MoveNext
                
            Loop
        
            If MsgBox("Confirm save employee's DTR?", vbQuestion + vbYesNo) = vbYes Then
            
                CitronPayroll.Execute "set autocommit = 0"
                CitronPayroll.BeginTrans
                CitronPayroll.Execute "delete from dtremp where empno = '" & tdbEmployee.BoundText & "' and " & _
                            "workdate between '" & Format(tdbPayrollPeriod.Columns("wrkdatefrom").Text, "YYYY-MM-DD") & "' and " & _
                            "'" & Format(tdbPayrollPeriod.Columns("wrkdateto").Text, "YYYY-MM-DD") & "'"
                .MoveFirst
                Do While Not .EOF
                
                    If Trim(!st1in) <> "" And Trim(!st1out) <> "" Then
                        If IsDate(!t1in) And Not IsDate(!t1out) Then
                            !t1in = ""
                        ElseIf Not IsDate(!t1in) And IsDate(!t1out) Then
                            !t1out = ""
                        End If
                    End If
                    
                    If Trim(!st2in) <> "" And Trim(!st2out) <> "" Then
                        If IsDate(!t2in) And Not IsDate(!t2out) Then
                            !t2in = ""
                        ElseIf Not IsDate(!t2in) And IsDate(!t2out) Then
                            !t2out = ""
                        End If
                    End If
                
                
                    CitronPayroll.Execute "insert into dtremp(empno,dayno,workdate,shiftcode,t1in,t1out,t2in,t2out, " & _
                            "st1in,st1out,st2in,st2out,brkstart,brkend,nitepremstart,nitepremend,wrkhrs,absent,latehrs,uthrs,dayoff,updatable,hrsperday,brkhrsperday) values " & _
                            "('" & tdbEmployee.BoundText & "','" & Weekday(!wrkdate) & "','" & Format(!wrkdate, "YYYY-MM-DD") & "','" & !shiftcode & "', " & _
                            "'" & IIf(Trim(!t1in) <> "", Format(!t1in, "hh:nn:ss"), "") & "','" & IIf(Trim(!t1out) <> "", Format(!t1out, "hh:nn:ss"), "") & "', " & _
                            "'" & IIf(Trim(!t2in) <> "", Format(!t2in, "hh:nn:ss"), "") & "','" & IIf(Trim(!t2out) <> "", Format(!t2out, "hh:nn:ss"), "") & "', " & _
                            "'" & IIf(Trim(!st1in) <> "", Format(!st1in, "hh:nn"), "") & "','" & IIf(Trim(!st1out) <> "", Format(!st1out, "hh:nn"), "") & "', " & _
                            "'" & IIf(Trim(!st2in) <> "", Format(!st2in, "hh:nn"), "") & "','" & IIf(Trim(!st2out) <> "", Format(!st2out, "hh:nn"), "") & "', " & _
                            "'" & IIf(Trim(!brkstart) <> "", Format(!brkstart, "hh:nn"), "") & "','" & IIf(Trim(!brkend) <> "", Format(!brkend, "hh:nn"), "") & "', " & _
                            "'" & IIf(Trim(!nitepremstart) <> "", Format(!nitepremstart, "hh:nn"), "") & "','" & IIf(Trim(!nitepremend) <> "", Format(!nitepremend, "hh:nn"), "") & "', " & _
                            "0,0,0,0,'" & IIf(!dayoff <> 0, "Y", "N") & "','" & IIf(!updatable <> 0, "Y", "N") & "', " & Format(!hrsperday, "###0.00") & ", " & Format(!brkhrsperday, "###0.00") & ")"
                    .MoveNext
                    DoEvents
                Loop
                .MoveLast
                CitronPayroll.CommitTrans
                tdgDtr.SetFocus
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
  
  bind_tdb CitronPayroll, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto from payrollperiod", "description", "percode"
  
  bind_tdb CitronPayroll, tdbEmployee, "select empno,concat(lastname, ', ', firstname,' ',middlename) fullname from employee order by concat(lastname, ', ', firstname,' ',middlename)", "fullname", "empno"
  
  Bind_tdd CitronPayroll, tddShift, "select *,concat(t1in,' ',t1out,'    ',t2in,' ',t2out) shiftdesc,(t1hrs+t2hrs) hrsperday, brkhrs brkhrsperday from shift", "shiftcode"
  
  CreateDtrTmp
  
  mPercode = ""
  mEmpNo = ""
  CreateDtrTmp
  cmdSave.Enabled = False
  
  Add_MDIButton Me.Name, TitleBar.Caption
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()
  
    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
      .Top = TitleBar.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With fra1
      .Top = frabutton.Top + frabutton.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With tdgDtr
      .Top = fra1.Height + fra1.Top
      .Left = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight - .Top
    End With

End Sub

Private Sub tdbPayrollPeriod_LostFocus()
  
  If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then
    If mPercode <> "" Then
      If mPercode <> tdbPayrollPeriod.BoundText Then
        If MsgBox("You selected another payroll period the system will now clear the grid." & vbCr & "Do you want to proceed?", vbInformation + vbYesNo) = vbYes Then
          mPercode = tdbPayrollPeriod.BoundText
          CreateDtrTmp
          cmdSave.Enabled = False
        Else
          tdbPayrollPeriod.BoundText = mPercode
        End If
      End If
    Else
      mPercode = tdbPayrollPeriod.BoundText
      CreateDtrTmp
    End If
  Else
    tdbPayrollPeriod.BoundText = mPercode
  End If
  
End Sub

Private Sub tdbPayrollPeriod_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
  End If
End Sub

Private Sub tdbEmployee_LostFocus()
  If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
    If mEmpNo <> "" Then
      If mEmpNo <> tdbEmployee.BoundText Then
        If MsgBox("You selected another employee the system will now clear the grid." & vbCr & "Do you want to proceed?", vbInformation + vbYesNo) = vbYes Then
          mEmpNo = tdbEmployee.BoundText
          CreateDtrTmp
          cmdSave.Enabled = False
        Else
          tdbEmployee.BoundText = mEmpNo
        End If
      End If
    Else
      mEmpNo = tdbEmployee.BoundText
      CreateDtrTmp
    End If
  Else
    tdbEmployee.BoundText = mEmpNo
  End If
End Sub

Private Sub tdbEmployee_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEmployee, tdbEmployee.RowSource, tdbEmployee.Text
  End If
End Sub

Private Sub CreateDtrTmp()
  
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
    .Fields.Append "shiftcode", adVarChar, 7
    .Fields.Append "shiftdetail", adVarChar, 50
    .Fields.Append "brkstart", adVarChar, 5
    .Fields.Append "brkend", adVarChar, 5
    .Fields.Append "nitepremstart", adVarChar, 5
    .Fields.Append "nitepremend", adVarChar, 5
    .Fields.Append "hrsperday", adDouble, 18
    .Fields.Append "brkhrsperday", adDouble, 18
    .Open
  End With
  
  Set tdgDtr.DataSource = rsDtrTmp
  
End Sub

Private Sub Load_Dtr()
  
    Dim mdate       As Date
    Dim rsEmpDtr    As ADODB.Recordset
    Dim rsEmpShift  As ADODB.Recordset
    Dim rsHoliday   As ADODB.Recordset
    Dim rsOBT       As ADODB.Recordset
    Dim rsLeave     As ADODB.Recordset
    
    CreateDtrTmp
    
    With rsDtrTmp
      
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
              
              NetOpen rsEmpShift, "", "select x2.*,(x2.t1hrs+x2.t2hrs) hrsperday, x2.brkhrs  from empshift x1 left outer join shift x2 on " & _
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
                    .Fields("brkhrsperday") = rsEmpShift!brkhrs
                    .Fields("dayoff") = 0
                    
              Else
                    .Fields("updatable") = 1
                    .Fields("dayoff") = 1
                    .Fields("hrsperday") = 0
                    .Fields("brkhrsperday") = 0
              End If
              
            End If
            
            NetOpen rsHoliday, "", "select * from holiday where holidaydate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
            If rsHoliday.RecordCount > 0 Then
              If CInt(rsHoliday!regular) = 1 Then
                .Fields("holiday") = "Legal"
              Else
                .Fields("holiday") = "Special"
              End If
            End If
            
            NetOpen rsOBT, "", "select x1.* from obtlne x1 left outer join obthdr x2 on x1.obtnum = x2.obtnum " & _
                             "where x2.empno = '" & tdbEmployee.BoundText & "' and x1.obtdate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
            If rsOBT.RecordCount > 0 Then
                .Fields("travel") = 1
            Else
                .Fields("travel") = 0
            End If
            
            NetOpen rsLeave, "", "select x1.* from lvlne x1 " & _
                    "left outer join lvhdr x2 on x1.lvnum = x2.lvnum " & _
                    "where x2.empno = '" & tdbEmployee.BoundText & "' and x1.lvdate = '" & Format(mdate, "YYYY-MM-DD") & "'"
            
            If rsLeave.RecordCount > 0 Then
                .Fields("leave") = 1
            Else
                .Fields("leave") = 0
            End If
            
            .Update
            
            mdate = mdate + 1
          
        Loop
      
    End With

End Sub

Private Sub tddShift_DropDownOpen()
  Bind_tdd CitronPayroll, tddShift, "select *,concat(t1in,'   ',t1out,'       ',t2in,'   ',t2out) shiftdesc, (t1hrs+t2hrs) hrsperday, brkhrs brkhrsperday from shift", "shiftcode"
End Sub

Private Sub tddShift_RowChange()
  With tdgDtr
    .Columns("shiftcode").Text = tddShift.Columns("shiftcode").Text
    txtShiftcode.Text = tddShift.Columns("shiftcode").Text
    .Columns("shiftdetail").Text = tddShift.Columns("shiftdesc").Text
    .Columns("st1in").Text = tddShift.Columns("t1in").Text
    .Columns("st1out").Text = tddShift.Columns("t1out").Text
    .Columns("st2in").Text = tddShift.Columns("t2in").Text
    .Columns("st2out").Text = tddShift.Columns("t2out").Text
    .Columns("brkstart").Text = tddShift.Columns("brkstart").Text
    .Columns("brkend").Text = tddShift.Columns("brkend").Text
    .Columns("nitepremstart").Text = tddShift.Columns("nitepremstart").Text
    .Columns("nitepremend").Text = tddShift.Columns("nitepremend").Text
    .Columns("hrsperday").Text = tddShift.Columns("hrsperday").Text
    .Columns("brkhrsperday").Text = tddShift.Columns("brkhrsperday").Text
  End With
End Sub

Private Sub tdgDtr_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = tdgDtr.Columns("dayoff").ColIndex Then
        If tdgDtr.Columns("dayoff").Value <> 0 Then
            tdgDtr.Columns(ColIndex).Value = 1
            tdgDtr.Columns("shiftcode").Text = ""
            tdgDtr.Columns("shiftdetail").Text = ""
            tdgDtr.Columns("t1in").Text = ""
            tdgDtr.Columns("t1out").Text = ""
            tdgDtr.Columns("t2in").Text = ""
            tdgDtr.Columns("t2out").Text = ""
            tdgDtr.Columns("st1in").Text = ""
            tdgDtr.Columns("st1out").Text = ""
            tdgDtr.Columns("st2in").Text = ""
            tdgDtr.Columns("st2out").Text = ""
            tdgDtr.Columns("brkstart").Text = ""
            tdgDtr.Columns("brkend").Text = ""
            tdgDtr.Columns("nitepremstart").Text = ""
            tdgDtr.Columns("nitepremend").Text = ""
            tdgDtr.Columns("hrsperday").Text = 0
            tdgDtr.Columns("brkhrsperday").Text = 0
        End If
    ElseIf ColIndex = tdgDtr.Columns("t1in").ColIndex Then
        If Not IsDate(tdgDtr.Columns("t1in").Text) Then
            tdgDtr.Columns("t1out").Text = ""
        End If
    ElseIf ColIndex = tdgDtr.Columns("t2in").ColIndex Then
        If Not IsDate(tdgDtr.Columns("t2in").Text) Then
            tdgDtr.Columns("t2out").Text = ""
        End If
    End If
End Sub

Private Sub tdgDtr_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = tdgDtr.Columns("shiftcode").ColIndex Then
        If tdgDtr.Columns("dayoff").Value <> 0 Then
            Cancel = True
        End If
    ElseIf ColIndex >= tdgDtr.Columns("t1in").ColIndex And ColIndex <= tdgDtr.Columns("t1out").ColIndex Then
        If tdgDtr.Columns("dayoff").Value <> 0 Then
            Cancel = True
            Exit Sub
        End If
        If Trim(tdgDtr.Columns("st1in").Text) = "" And Trim(tdgDtr.Columns("st1out").Text) = "" Then
            Cancel = True
            Exit Sub
        End If
        If ColIndex = tdgDtr.Columns("t1out").ColIndex Then
            If Trim(tdgDtr.Columns("t1in").Text) = "" Then
                Cancel = True
                Exit Sub
            End If
        End If
    ElseIf ColIndex >= tdgDtr.Columns("t2in").ColIndex And ColIndex <= tdgDtr.Columns("t2out").ColIndex Then
        If tdgDtr.Columns("dayoff").Value <> 0 Then
            Cancel = True
            Exit Sub
        End If
        If Trim(tdgDtr.Columns("st2in").Text) = "" And Trim(tdgDtr.Columns("st2out").Text) = "" Then
            Cancel = True
            Exit Sub
        End If
        If ColIndex = tdgDtr.Columns("t2out").ColIndex Then
            If Trim(tdgDtr.Columns("t2in").Text) = "" Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub tdgDtr_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If ColIndex >= tdgDtr.Columns("t1in").ColIndex And ColIndex <= tdgDtr.Columns("t2out").ColIndex Then
        If Not IsDate(tdgDtr.Columns(ColIndex).Text) Then
            txtTime.Text = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub tdgDtr_BeforeRowColChange(Cancel As Integer)
    
    With tdgDtr
        If Trim(.Columns("st1in").Text) <> "" And Trim(.Columns("st1out").Text) <> "" And Trim(.Columns("st2in").Text) = "" And Trim(.Columns("st2out").Text) = "" Then
            If Trim(.Columns("t1in").Text) <> "" And Trim(.Columns("t1out").Text) = "" Then
                .Col = .Columns("t1out").ColIndex
                Cancel = True
            End If
        ElseIf Trim(.Columns("st1in").Text) <> "" And Trim(.Columns("st1out").Text) <> "" And Trim(.Columns("st2in").Text) <> "" And Trim(.Columns("st2out").Text) <> "" Then
            If Trim(.Columns("t1in").Text) <> "" And Trim(.Columns("t1out").Text) = "" Then
                .Col = .Columns("t1out").ColIndex
                Cancel = True
            ElseIf Trim(.Columns("t2in").Text) <> "" And Trim(.Columns("t2out").Text) = "" Then
                .Col = .Columns("t2out").ColIndex
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub txtDayStat_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tdgDtr.SetFocus
    Else
        If KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") And KeyAscii <> Asc("y") And KeyAscii <> Asc("n") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDayStat_Change()
    txtDayStat.Text = UCase(txtDayStat.Text)
End Sub

Private Sub txtDayStat_LostFocus()
    tdgDtr.SetFocus
End Sub

Private Sub txtshiftcode_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tdgDtr.SetFocus
  Else
    SearchRecord KeyAscii, txtShiftcode, tddShift.DataSource, txtShiftcode.Text, "shiftcode"
    With tdgDtr
      .Columns("shiftdetail").Text = tddShift.Columns("shiftdesc").Text
      .Columns("st1in").Text = tddShift.Columns("t1in").Text
      .Columns("st1out").Text = tddShift.Columns("t1out").Text
      .Columns("st2in").Text = tddShift.Columns("t2in").Text
      .Columns("st2out").Text = tddShift.Columns("t2out").Text
      .Columns("brkstart").Text = tddShift.Columns("brkstart").Text
      .Columns("brkend").Text = tddShift.Columns("brkend").Text
      .Columns("nitepremstart").Text = tddShift.Columns("nitepremstart").Text
      .Columns("nitepremend").Text = tddShift.Columns("nitepremend").Text
      .Columns("hrsperday").Text = tddShift.Columns("hrsperday").Text
      .Columns("brkhrsperday").Text = tddShift.Columns("brkhrsperday").Text
    End With
  End If
End Sub

Private Sub txtTime_LostFocus()
  tdgDtr.SetFocus
End Sub
