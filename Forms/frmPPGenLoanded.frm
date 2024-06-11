VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPGenLoanDed 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11685
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPPGenLoanded.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   11685
   Tag             =   "Loan Deductions"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSearch 
      BackColor       =   &H00808080&
      ForeColor       =   &H00404040&
      Height          =   1005
      Left            =   75
      TabIndex        =   10
      Top             =   780
      Width           =   13620
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   6315
         TabIndex        =   2
         Top             =   615
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   529
         Caption         =   "frmPPGenLoanded.frx":6852
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPPGenLoanded.frx":68BE
         Key             =   "frmPPGenLoanded.frx":68DC
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
         Left            =   1650
         TabIndex        =   1
         Tag             =   "Municipal"
         Top             =   585
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
         _PropDict       =   $"frmPPGenLoanded.frx":6920
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
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1650
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
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "payfreqcode"
         Columns(4).DataField=   "payfreqcode"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "paymonth"
         Columns(5).DataField=   "paymonth"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "payyear"
         Columns(6).DataField=   "payyear"
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
         Splits(0)._ColumnProps(22)=   "Column(4).Width=3254"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=3254"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=3254"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=3175"
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
         _PropDict       =   $"frmPPGenLoanded.frx":69CA
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
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "PAYROLL PERIOD"
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
         Height          =   225
         Left            =   195
         TabIndex        =   13
         Top             =   240
         Width           =   1830
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
         Left            =   5250
         TabIndex        =   12
         Top             =   645
         Width           =   1005
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
         Left            =   585
         TabIndex        =   11
         Top             =   630
         Width           =   1005
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
      ScaleWidth      =   11685
      TabIndex        =   8
      Top             =   0
      Width           =   11685
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Deductions"
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
         Tag             =   "Loan Deductions"
         Top             =   225
         Width           =   4410
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   30
      TabIndex        =   7
      Top             =   9285
      Width           =   7410
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   390
         Left            =   45
         TabIndex        =   5
         Top             =   15
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
         Image           =   "frmPPGenLoanded.frx":6A74
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   390
         Left            =   2055
         TabIndex        =   4
         Top             =   15
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
         Image           =   "frmPPGenLoanded.frx":774E
         cBack           =   14737632
      End
   End
   Begin VB.Frame fraLoanList 
      BackColor       =   &H00E0E0E0&
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
      Height          =   6390
      Left            =   150
      TabIndex        =   6
      Top             =   1980
      Width           =   14805
      Begin TrueOleDBGrid80.TDBGrid tdgGenLoanDed 
         Height          =   3000
         Left            =   30
         TabIndex        =   3
         Top             =   150
         Width           =   14580
         _ExtentX        =   25718
         _ExtentY        =   5292
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Employee"
         Columns(0).DataField=   "fullname"
         Columns(0).DropDown=   "tddChargeName"
         Columns(0).DropDown.vt=   8
         Columns(0).ExternalEditor=   "txtChargeName"
         Columns(0).ExternalEditor.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Loan Types"
         Columns(1).DataField=   "loantypesname"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Amount Deducted"
         Columns(2).DataField=   "amtded"
         Columns(2).NumberFormat=   "#,##0.00"
         Columns(2).ExternalEditor=   "txtamtded"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "balance"
         Columns(4).DataField=   "balance"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ttlamtpaid"
         Columns(5).DataField=   "ttlamtpaid"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=7091"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=7011"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6588"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6509"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8704"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=5794"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerStyle=0"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=5741"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(2)._HeadDivider=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=8705"
         Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(35)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=8708"
         Splits(0)._ColumnProps(40)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=86,.parent=13,.alignment=1,.locked=0"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
      Begin TDBNumber6Ctl.TDBNumber txtamtded 
         Height          =   270
         Left            =   90
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   476
         Calculator      =   "frmPPGenLoanded.frx":8428
         Caption         =   "frmPPGenLoanded.frx":8448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPPGenLoanded.frx":84B4
         Keys            =   "frmPPGenLoanded.frx":84D2
         Spin            =   "frmPPGenLoanded.frx":851C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
         MaxValue        =   999999999
         MinValue        =   -999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
   End
End
Attribute VB_Name = "frmPPGenLoanDed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsLoanDed       As ADODB.Recordset
Dim rsTmpLoanded    As ADODB.Recordset

Private Sub GenerateLoan()

    Dim rsPrevPeriod  As ADODB.Recordset
    Dim rsBalance     As ADODB.Recordset

    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        
        MsgBox "Please choose a payroll period before generating a loan deduction.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
        
    End If
    
    NetOpen rsPrevPeriod, "SELECT COUNT(percode) ctr FROM payrollperiod WHERE percode < " & tdbPayrollPeriod.BoundText & " AND fnlz = 'N'"
    
    If rsPrevPeriod.RecordCount > 0 Then
      If rsPrevPeriod.Fields("ctr") > 0 Then
        MsgBox "Please finalize the previous payroll before continuing to generate the current loan deductions.", vbExclamation + vbOKOnly
        Exit Sub
      End If
    End If
    
    Create_TmpLoanDed
    
    With rsTmpLoanded
    
        NetOpen rsLoanDed, "select x1.*, concat(x2.lastname,', ',x2.firstname,' ',x2.middlename ) fullname,x3.loantypesname from loans x1 " & _
                               "left outer join employee x2 on x1.employeecode = x2.employeecode " & _
                               "left outer join loantypes x3 on x1.loantypescode = x3.loantypescode " & _
                               "where x1.status = 'Active' and x1.startdate <= '" & Format(tdbPayrollPeriod.Columns("to").Text, "YYYY-MM-DD") & "' and " & _
                               "x2.isactive = 'Y' and x2.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "' and " & _
                               " x1.loantypescode in (select loantypescode from payrollperiodloandedallow where allow = 1 and percode = " & tdbPayrollPeriod.BoundText & ") " & _
                               "order by x2.lastname,x2.firstname,x2.middlename "
        
        If rsLoanDed.RecordCount > 0 Then
            
            rsLoanDed.MoveFirst
            
            Do While Not rsLoanDed.EOF
                DoEvents
                .AddNew
                .Fields("employeecode") = rsLoanDed!employeecode
                .Fields("fullname") = rsLoanDed!fullname
                .Fields("loancode") = rsLoanDed!loancode
                .Fields("loantypescode") = rsLoanDed!loantypescode
                .Fields("loantypesname") = rsLoanDed!loantypesname
                
                NetOpen rsBalance, "select balance,ttlamtpaid from loanded " & _
                                       "where loancode = '" & rsLoanDed!loancode & "'  and fnlz = 'Y' " & _
                                       "order by loandedcode desc limit 1"
                                       
                If CDbl(rsLoanDed!dedperpayday) > rsBalance!balance Then
                    .Fields("amtded") = rsBalance!balance
                Else
                    If (rsBalance!balance - rsLoanDed!dedperpayday) / rsLoanDed!dedperpayday < 0.1 Then
                        .Fields("amtded") = rsBalance!balance
                        .Fields("dedperpayday") = rsBalance!balance
                    Else
                        .Fields("amtded") = rsLoanDed!dedperpayday
                        .Fields("dedperpayday") = rsLoanDed!dedperpayday
                    End If
                End If
                
                .Fields("balance") = rsBalance!balance
                .Fields("ttlamtpaid") = rsBalance!ttlamtpaid
                
                .Update
                
                rsLoanDed.MoveNext
                
                DoEvents
                
            Loop
        End If
    
        NetOpen rsLoanDed, "select * from loanded where percode = '" & tdbPayrollPeriod.BoundText & "'"
        
        If rsLoanDed.RecordCount > 0 Then
            rsLoanDed.MoveFirst
            Do While Not rsLoanDed.EOF
                .MoveFirst
                .Find "loancode = '" & rsLoanDed!loancode & "'"
                If Not .EOF Then
                    .Fields("amtded") = rsLoanDed!amtded
                End If
                rsLoanDed.MoveNext
                DoEvents
            Loop
        End If
        
    End With
    
    If rsTmpLoanded.RecordCount > 0 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
        MsgBox "No record found.", vbExclamation + vbOKOnly
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    Dim rsDateTime          As ADODB.Recordset
    Dim mBalance            As Double
    Dim mNoOfPay            As Integer
    
    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"
    
    With rsTmpLoanded
        If .RecordCount > 0 Then
            If MsgBox("Confirm save loan deductions entries.", vbInformation + vbOKCancel) = vbOK Then
                
                .MoveFirst
                
                ConMain.Execute "set autocommit = 0"
                ConMain.BeginTrans
                ConMain.Execute "delete from loanded where percode = '" & tdbPayrollPeriod.BoundText & "'"
                
                Do While Not .EOF
                    If !amtded <> 0 Then
                        
                        mBalance = CDbl(!balance) - CDbl(!amtded)
                        mNoOfPay = 0
                        
                        If !employeecode = 23 And !loantypescode = 1 Then
                            mNoOfPay = 0
                        End If
                        If mBalance > 0 And !dedperpayday > 0 Then
                            mNoOfPay = Compute_Inst(mBalance, !dedperpayday)
                        End If
                        
                        ConMain.Execute "update payroll set fnlz = 'N' where percode = " & tdbPayrollPeriod.BoundText & " and employeecode = " & !employeecode & ""
                        ConMain.Execute "insert into loanded(loancode,loantypescode,employeecode,percode,payyear,paymonth, " & _
                                        "amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled,usercode,noofpay) values " & _
                                        "('" & !loancode & "','" & !loantypescode & "','" & !employeecode & "','" & tdbPayrollPeriod.BoundText & "','" & tdbPayrollPeriod.Columns("payyear") & "','" & tdbPayrollPeriod.Columns("paymonth") & "', " & _
                                        "" & !amtded & ",'" & Format(rsDateTime!currentdate, "YYYY-MM-DD") & "'," & CDbl(!ttlamtpaid) + CDbl(!amtded) & "," & CDbl(!balance) - CDbl(!amtded) & ",'N','N'," & GlobalUserID & "," & mNoOfPay & ")"
                    
                    End If
                    .MoveNext
                Loop
                
                ConMain.Execute "update payrollperiod set genpay = 'N' where percode = " & tdbPayrollPeriod.BoundText & ""
                
                ConMain.CommitTrans
                
                MsgBox "Process complete!", vbInformation + vbOKOnly
                
            End If
        End If
    End With
    
    tdgGenLoanDed.SetFocus
    
End Sub

Private Sub Form_Activate()

    Focus_MDIButton Me
    
End Sub

Private Sub Form_Load()
    
    Dim rsTmp           As ADODB.Recordset
    
    Dim i               As Integer
    
    Add_MDIButton Me.Name, Me.Tag
    
    bind_tdb ConMain, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto,payfreqcode,paymonth,payyear from payrollperiod where fnlz <> 'Y' order by percode desc", "description", "percode"
    
    Create_TmpLoanDed
    
    cmdSave.Enabled = False
    
    CreateTmpDB rsTmp
    
    With tdgGenLoanDed
        For i = .Columns("fullname").ColIndex To .Columns("loantypesname").ColIndex
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
        .BoundText = "fullname"
    End With
    
    Set rsTmp = Nothing
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
    Set rsLoanDed = Nothing
    Set rsTmpLoanded = Nothing

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With fraSearch
        .Top = pic1.Top + pic1.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fraLoanList
        .Top = fraSearch.Top + fraSearch.Height
        .Left = 0
        .Height = Me.ScaleHeight - (.Top + fraButtons.Height)
        .Width = Me.ScaleWidth
    End With
    
    With tdgGenLoanDed
        .Top = 200
        .Left = 50
        .Height = fraLoanList.Height - 300
        .Width = Me.ScaleWidth - 100
    End With
    
'    With fraButton1
'        .Top = tdgGenloanded.Top + tdgGenloanded.Height + 50
'        .Left = tdgGenloanded.Left
'        .Width = tdgGenloanded.Width
'    End With
    
    With fraButtons
        .Top = fraLoanList.Top + fraLoanList.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With

End Sub

Private Sub Create_TmpLoanDed()

    Set rsTmpLoanded = Nothing
    Set rsTmpLoanded = New ADODB.Recordset
    
    With rsTmpLoanded
        .Fields.Append "employeecode", adVarChar, 15
        .Fields.Append "fullname", adVarChar, 70
        .Fields.Append "loancode", adVarChar, 7
        .Fields.Append "loantypescode", adVarChar, 7
        .Fields.Append "loantypesname", adVarChar, 70
        .Fields.Append "amtded", adDouble
        .Fields.Append "balance", adDouble
        .Fields.Append "dedperpayday", adDouble
        .Fields.Append "ttlamtpaid", adDouble
        .Open
        .Sort = "fullname"
    End With
        
    Set tdgGenLoanDed.DataSource = rsTmpLoanded
    
End Sub

Private Sub tdbPayrollPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        DoEvents
        GenerateLoan
    Else
        SearchList KeyAscii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
    End If
End Sub

Private Sub tdbSearch_ItemChange()
    rsTmpLoanded.Sort = tdbSearch.BoundText
End Sub

Private Sub tdbSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbSearch, tdbSearch.RowSource, tdbSearch.Text
        tdbSearch_ItemChange
    End If
End Sub

Private Sub tdgGenLoanDed_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    With tdgGenLoanDed
        If ColIndex = .Columns("amtded").ColIndex Then
            If Not IsNumeric(.Columns("amtded").Text) Then
                .Columns("amtded").Text = "0.00"
            End If
            If CDbl(.Columns("amtded").Text) > CDbl(.Columns("balance").Text) Then
                MsgBox "The amount you entered is greater than the remaining balance of the employee's loan.", vbExclamation + vbOKOnly
                Cancel = True
                .SetFocus
                .Col = .Columns("amtded").ColIndex
            End If
        End If
    End With
End Sub


Private Sub tdgGenLoanDed_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
End Sub

Private Sub tdgGenLoanDed_KeyPress(KeyAscii As Integer)
    With tdgGenLoanDed
        If KeyAscii = 13 Then
            If .Col - 1 = .Columns("amtded").ColIndex Then
                If .Row < .VisibleRows - 2 Then
                    .Row = .Row + 1
                    .Col = .Columns("amtded").ColIndex
                ElseIf .Row = .VisibleRows - 2 Then
                    SendKeys "{DOWN}"
                    .Col = .Columns("amtded").ColIndex
                ElseIf .Row > .ApproxCount - 1 Then
                    .Col = .Columns("amtded").ColIndex
                    SendKeys "{TAB}"
                End If
            End If
        ElseIf KeyAscii = 6 Then
            txtSearch.SetFocus
        End If
    End With
End Sub

Private Sub txtamtded_LostFocus()
    On Error Resume Next
    tdgGenLoanDed.SetFocus
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(txtSearch.Text)
    End With
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With tdgGenLoanDed
            .SetFocus
            If rsTmpLoanded.RecordCount > 0 Then
                If Not rsTmpLoanded.EOF Then
                    .Col = .Columns("amtded").ColIndex
                End If
            End If
        End With
    Else
        SearchRecord KeyAscii, txtSearch, rsTmpLoanded, txtSearch.Text, tdbSearch.BoundText
    End If
End Sub

Private Function Compute_Inst(mBal As Double, mDed As Double) As Integer
  
    Dim mInst         As Double
    Dim i             As Integer
    
    i = 0
        
    mInst = CDbl(mBal)
    
    If mInst >= mDed Then
        Do While mInst > 0
            
            If mInst >= CDbl(mDed) Then
                mInst = mInst - CDbl(mDed)
                i = i + 1
            Else
                Exit Do
            End If
            
        Loop
    Else
        i = 1
    End If
    
    If i > 1 Then
        If mInst > 0 Then
            If (mInst / CDbl(mDed)) > 0.1 Then
                i = i + 1
            End If
        End If
    End If
    

     Compute_Inst = i
     
End Function

