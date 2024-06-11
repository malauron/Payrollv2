VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADOvertime3 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CitronSoftwarePayroll.b8ChildTitleBar b8ChildTitleBar1 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Overtime Detail"
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
      ForeColor       =   16777215
      GradTheme       =   1
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   0
      TabIndex        =   1
      Top             =   255
      Width           =   11265
      Begin VB.Frame fra1 
         BackColor       =   &H00FFFFFF&
         Height          =   3720
         Left            =   90
         TabIndex        =   12
         Top             =   105
         Width           =   4905
         Begin TDBText6Ctl.TDBText txtCostCenter 
            Height          =   300
            Left            =   1665
            TabIndex        =   13
            Top             =   1770
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":006C
            Key             =   "frmADOvertime3.frx":008A
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
         Begin TDBText6Ctl.TDBText txtDivision 
            Height          =   300
            Left            =   1665
            TabIndex        =   14
            Top             =   2100
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":00CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":013A
            Key             =   "frmADOvertime3.frx":0158
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
         Begin TDBText6Ctl.TDBText txtBranch 
            Height          =   300
            Left            =   1665
            TabIndex        =   15
            Top             =   2430
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":019C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0208
            Key             =   "frmADOvertime3.frx":0226
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
         Begin TDBText6Ctl.TDBText txtApprovBy 
            Height          =   300
            Left            =   1665
            TabIndex        =   16
            Top             =   3195
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":026A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":02D6
            Key             =   "frmADOvertime3.frx":02F4
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
         Begin TrueOleDBList80.TDBCombo tdbEmpNo 
            Height          =   315
            Left            =   1665
            TabIndex        =   17
            Tag             =   "Municipal"
            Top             =   585
            Width           =   3000
            _ExtentX        =   5292
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
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
            _PropDict       =   $"frmADOvertime3.frx":0338
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
         Begin TrueOleDBList80.TDBCombo tdbEmpName 
            Height          =   345
            Left            =   1665
            TabIndex        =   18
            Tag             =   "Municipal"
            Top             =   960
            Width           =   3000
            _ExtentX        =   5292
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
            Columns(2).Caption=   "costcentercode"
            Columns(2).DataField=   "costcentercode"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "divisioncode"
            Columns(3).DataField=   "divisioncode"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "branchcode"
            Columns(4).DataField=   "branchcode"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "costcenter"
            Columns(5).DataField=   "costcenter"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "division"
            Columns(6).DataField=   "division"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "branch"
            Columns(7).DataField=   "branch"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
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
            _PropDict       =   $"frmADOvertime3.frx":03E2
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
            _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText txtOTCode 
            Height          =   300
            Left            =   1665
            TabIndex        =   19
            Top             =   255
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":048C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":04F8
            Key             =   "frmADOvertime3.frx":0516
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
         Begin VB.Line Line1 
            BorderColor     =   &H0030A0B8&
            X1              =   75
            X2              =   4770
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0030A0B8&
            X1              =   75
            X2              =   4770
            Y1              =   3030
            Y2              =   3030
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Number"
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
            Height          =   240
            Left            =   165
            TabIndex        =   26
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Employee ID"
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
            Height          =   240
            Left            =   165
            TabIndex        =   25
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Employee Name"
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
            Height          =   240
            Left            =   165
            TabIndex        =   24
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   240
            Left            =   150
            TabIndex        =   23
            Top             =   1785
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   240
            Left            =   135
            TabIndex        =   22
            Top             =   2145
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   240
            Left            =   135
            TabIndex        =   21
            Top             =   2505
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Approved By"
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
            Height          =   300
            Left            =   45
            TabIndex        =   20
            Top             =   3240
            Width           =   1545
         End
      End
      Begin VB.Frame fra3 
         BackColor       =   &H00FFFFFF&
         Height          =   2340
         Left            =   5040
         TabIndex        =   5
         Top             =   105
         Width           =   6150
         Begin TDBText6Ctl.TDBText txtReason 
            Height          =   300
            Left            =   1680
            TabIndex        =   6
            Top             =   1890
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":055A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":05C6
            Key             =   "frmADOvertime3.frx":05E4
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
            Height          =   300
            Left            =   1680
            TabIndex        =   7
            Top             =   570
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":0628
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0694
            Key             =   "frmADOvertime3.frx":06B2
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
         Begin TDBText6Ctl.TDBText txtSchedule 
            Height          =   300
            Left            =   1680
            TabIndex        =   8
            Top             =   900
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":06F6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0762
            Key             =   "frmADOvertime3.frx":0780
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
         Begin TDBText6Ctl.TDBText txtActTito 
            Height          =   300
            Left            =   1680
            TabIndex        =   9
            Top             =   1230
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":07C4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0830
            Key             =   "frmADOvertime3.frx":084E
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
         Begin TDBDate6Ctl.TDBDate txtApprovdate 
            Height          =   300
            Left            =   1680
            TabIndex        =   10
            Top             =   1560
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   529
            Calendar        =   "frmADOvertime3.frx":0892
            Caption         =   "frmADOvertime3.frx":0998
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":09FE
            Keys            =   "frmADOvertime3.frx":0A1C
            Spin            =   "frmADOvertime3.frx":0A7A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            Text            =   "01/21/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39468
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate txtWorkdate 
            Height          =   300
            Left            =   1695
            TabIndex        =   11
            Top             =   240
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   529
            Calendar        =   "frmADOvertime3.frx":0AA2
            Caption         =   "frmADOvertime3.frx":0BA8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0C0E
            Keys            =   "frmADOvertime3.frx":0C2C
            Spin            =   "frmADOvertime3.frx":0C8A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            Text            =   "01/21/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39468
            CenturyMode     =   0
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reason/Purpose"
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
            Height          =   240
            Left            =   150
            TabIndex        =   38
            Top             =   1935
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date Approved"
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
            Height          =   240
            Left            =   150
            TabIndex        =   37
            Top             =   1605
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Actal TITO"
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
            Height          =   240
            Left            =   150
            TabIndex        =   36
            Top             =   1290
            Width           =   1455
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Shift Schedule"
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
            Height          =   240
            Left            =   165
            TabIndex        =   35
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Day Status"
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
            Height          =   240
            Left            =   165
            TabIndex        =   34
            Top             =   615
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Work Date"
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
            Height          =   240
            Left            =   165
            TabIndex        =   33
            Top             =   285
            Width           =   1455
         End
      End
      Begin VB.Frame fraButton 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   75
         TabIndex        =   2
         Top             =   3795
         Width           =   3645
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   330
            Left            =   30
            TabIndex        =   3
            Top             =   60
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
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   330
            Left            =   1815
            TabIndex        =   4
            Top             =   60
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
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
            cFore           =   3186872
            cFHover         =   3186872
            cBhover         =   16777215
            cGradient       =   16777215
            Gradient        =   4
            Mode            =   0
            Value           =   0   'False
            cBack           =   14737632
         End
      End
      Begin VB.Frame fra2 
         BackColor       =   &H00FFFFFF&
         Height          =   1425
         Left            =   5040
         TabIndex        =   27
         Top             =   2400
         Width           =   6150
         Begin TDBNumber6Ctl.TDBNumber txtGrossOtAmnt 
            Height          =   315
            Left            =   4545
            TabIndex        =   28
            Top             =   600
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   556
            Calculator      =   "frmADOvertime3.frx":0CB2
            Caption         =   "frmADOvertime3.frx":0CD2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0D3E
            Keys            =   "frmADOvertime3.frx":0D5C
            Spin            =   "frmADOvertime3.frx":0DA6
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
            ForeColor       =   -2147483640
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtOTHrs 
            Height          =   315
            Left            =   1170
            TabIndex        =   29
            Top             =   930
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   556
            Calculator      =   "frmADOvertime3.frx":0DCE
            Caption         =   "frmADOvertime3.frx":0DEE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0E5A
            Keys            =   "frmADOvertime3.frx":0E78
            Spin            =   "frmADOvertime3.frx":0EC2
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
            ForeColor       =   -2147483640
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
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtHrlyRate 
            Height          =   315
            Left            =   4545
            TabIndex        =   30
            Top             =   255
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   556
            Calculator      =   "frmADOvertime3.frx":0EEA
            Caption         =   "frmADOvertime3.frx":0F0A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmADOvertime3.frx":0F76
            Keys            =   "frmADOvertime3.frx":0F94
            Spin            =   "frmADOvertime3.frx":0FDE
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
            ForeColor       =   -2147483640
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBTime6Ctl.TDBTime txtOTStart 
            Height          =   300
            Left            =   1170
            TabIndex        =   31
            Top             =   270
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":1006
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmADOvertime3.frx":1072
            Spin            =   "frmADOvertime3.frx":10C2
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtOTEnd 
            Height          =   300
            Left            =   1170
            TabIndex        =   32
            Top             =   600
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   529
            Caption         =   "frmADOvertime3.frx":10EA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmADOvertime3.frx":1156
            Spin            =   "frmADOvertime3.frx":11A6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gross OT Amount"
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
            Height          =   240
            Left            =   2790
            TabIndex        =   43
            Top             =   645
            Width           =   1725
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hourly Rate"
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
            Height          =   240
            Left            =   3465
            TabIndex        =   42
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "OT Hours"
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
            Height          =   240
            Left            =   90
            TabIndex        =   41
            Top             =   975
            Width           =   1035
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "OT End"
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
            Height          =   240
            Left            =   90
            TabIndex        =   40
            Top             =   645
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "OT Start"
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
            Height          =   240
            Left            =   90
            TabIndex        =   39
            Top             =   330
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frmADOvertime3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd         As Boolean

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  
  If Trim(tdbEmpName.Text) = "" Or IsNull(tdbEmpName.SelectedItem) Or tdbEmpName.ApproxCount = 0 Then
    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
    tdbEmpName.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(txtWorkdate.Text) Then
    MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
    txtWorkdate.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(txtApprovdate.Text) Then
    MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
    txtApprovdate.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(txtOTStart.Text) Then
    MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
    txtOTStart.SetFocus
    Exit Sub
  End If
  
  If Not IsDate(txtOTEnd.Text) Then
    MsgBox "Please enter valid time.", vbExclamation + vbOKOnly
    txtOTEnd.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(txtHrlyRate.Text) Then
    MsgBox "Please enter a valid number.", vbExclamation + vbOKOnly
    txtHrlyRate.SetFocus
    Exit Sub
  End If
  
  If Not IsNumeric(txtGrossOtAmnt.Text) Then
    MsgBox "Please enter a vaild number.", vbExclamation + vbOKOnly
    txtGrossOtAmnt.SetFocus
    Exit Sub
  End If
  
  
  If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  If mAdd = True Then
  
    CitronPayroll.Execute "set autocommit = 0"
    CitronPayroll.BeginTrans
    txtOTCode.Text = LastCode("GetLastCodeA", "Overtime", "0000000")
    
    CitronPayroll.Execute "insert into overtime (otcode,empno,percode,costcentercode,divisioncode,branchcode,wrkdate,approvby,approvdate,remarks, " & _
                          "otstart,otend,othrs,hrlyrate,grossotamnt,status,payyear,paymonth,payfreqcode) values ('" & txtOTCode.Text & "','" & tdbEmpName.BoundText & "', " & _
                          "'" & frmADOvertime.tdbPayrollPeriod.BoundText & "','" & tdbEmpName.Columns("costcentercode").Text & "','" & tdbEmpName.Columns("divisioncode").Text & "', " & _
                          "'" & tdbEmpName.Columns("branchcode").Text & "','" & Format(txtWorkdate.Text, "YYYY-MM-DD") & "','" & txtApprovBy.Text & "','" & Format(txtApprovdate.Text, "YYYY-MM-DD") & "', " & _
                          "'" & txtReason.Text & "','" & Format(txtOTStart.Text, "hh:nn") & "','" & Format(txtOTEnd.Text, "hh:nn") & "', " & _
                          "" & txtOTHrs.Text & "," & txtHrlyRate.Text & "," & txtGrossOtAmnt.Text & ",'Approved', " & _
                          "'" & frmADOvertime.tdbPayrollPeriod.Columns("payyear").Text & "','" & frmADOvertime.tdbPayrollPeriod.Columns("paymonth").Text & "','" & frmADOvertime.tdbPayrollPeriod.Columns("payfreqcode").Text & "')"
  
    CitronPayroll.CommitTrans
    
  Else
  
    CitronPayroll.Execute "set autocommit = 0"
    CitronPayroll.BeginTrans
    
    CitronPayroll.Execute "update overtime set empno = '" & tdbEmpName.BoundText & "', " & _
                          "costcentercode = '" & tdbEmpName.Columns("costcentercode").Text & "', divisioncode = '" & tdbEmpName.Columns("divisioncode").Text & "', " & _
                          "branchcode = '" & tdbEmpName.Columns("branchcode").Text & "',wrkdate = '" & Format(txtWorkdate.Text, "YYYY-MM-DD") & "',approvby = '" & txtApprovBy.Text & "', approvdate = '" & Format(txtApprovdate.Text, "YYYY-MM-DD") & "', " & _
                          "percode = '" & frmADOvertime.tdbPayrollPeriod.BoundText & "',remarks = '" & txtReason.Text & "', otstart = '" & Format(txtOTStart.Text, "hh:nn") & "', otend = '" & Format(txtOTEnd.Text, "hh:nn") & "', " & _
                          "othrs = " & txtOTHrs.Text & ", hrlyrate = " & txtHrlyRate.Text & ", grossotamnt = " & txtGrossOtAmnt.Text & ", " & _
                          "payyear = '" & frmADOvertime.tdbPayrollPeriod.Columns("payyear").Text & "', paymonth = '" & frmADOvertime.tdbPayrollPeriod.Columns("paymonth").Text & "', " & _
                          "payfreqcode = '" & frmADOvertime.tdbPayrollPeriod.Columns("payfreqcode").Text & "'  where otcode = '" & txtOTCode.Text & "'"
  
    CitronPayroll.CommitTrans
    
  End If
  
  frmADOvertime.rsOvertime.Requery
  frmADOvertime.rsOvertime.Find "otcode = '" & txtOTCode.Text & "'"
  Lock_Button "TTFFTT", frmADOvertime.cmdMenu, 5
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  
  bind_tdb CitronPayroll, tdbEmpNo, "select empno,empno from employee order by empno", "empno", "empno"
  
  bind_tdb CitronPayroll, tdbEmpName, "select x1.empno,concat(x1.lastname,' ',x1.firstname,' ',x1.middlename) fullname, " & _
                          "x1.costcentercode, x1.divisioncode, x1.branchcode,x2.costcenter,x3.division,x4.branch from employee x1 " & _
                          "left outer join costcenter x2 on x1.costcentercode = x2.costcentercode " & _
                          "left outer join division x3 on x1.divisioncode = x3.divisioncode " & _
                          "left outer join branch x4 on x1.branchcode = x4.branchcode order by concat(x1.lastname,' ',x1.firstname,' ',x1.middlename)", "fullname", "empno"
                          
  
  If mAdd = True Then
  
    txtWorkdate.Text = ""
    txtApprovdate.Text = ""
    txtOTStart.Text = ""
    txtOTEnd.Text = ""
    
  Else
  
    With frmADOvertime.rsOvertime
      If .RecordCount > 0 Then
        txtOTCode.Text = !otcode
        tdbEmpNo.BoundText = !empno
        tdbEmpName.BoundText = !empno
        tdbEmpNo.Enabled = False
        tdbEmpName.Enabled = False
        txtCostCenter.Text = !CostCenter & ""
        txtDivision.Text = !Division & ""
        txtBranch.Text = !branch & ""
        txtApprovBy.Text = !approvdby & ""
        txtWorkdate.Text = Format(!wrkdate, "MM/DD/YYYY")
        txtApprovdate.Text = Format(!approvdate, "MM/DD/YYYY")
        txtReason.Text = !remarks
        txtOTStart.Text = Format(!otstart, "hh:nn am/pm")
        txtOTEnd.Text = Format(!otend, "hh:nn am/pm")
        txtHrlyRate.Text = Format(!hrlyrate, "#,##0.00")
        txtGrossOtAmnt.Text = Format(!grossotamnt, "#,##0.00")
        txtOTHrs.Text = Format(!othrs, "#,##0.00")
        DoEvents
        Get_Sched !empno
      End If
    End With
    
  End If
  
End Sub

Private Sub tdbEmpName_Keypress(Keyascii As Integer)
  If Keyascii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList Keyascii, tdbEmpName, tdbEmpName.RowSource, tdbEmpName.Text
  End If
End Sub

Private Sub tdbEmpName_LostFocus()
  If tdbEmpNo.ApproxCount > 0 Then
    tdbEmpNo.BoundText = tdbEmpName.BoundText
    txtCostCenter.Text = tdbEmpName.Columns("costcenter").Text
    txtDivision.Text = tdbEmpName.Columns("division").Text
    txtBranch.Text = tdbEmpName.Columns("branch").Text
  End If
End Sub

Private Sub tdbEmpNo_Keypress(Keyascii As Integer)
  If Keyascii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList Keyascii, tdbEmpNo, tdbEmpNo.RowSource, tdbEmpNo.Text
    tdbEmpNo_LostFocus
  End If
End Sub

Private Sub tdbEmpNo_LostFocus()
  If tdbEmpName.ApproxCount > 0 Then
    tdbEmpName.BoundText = tdbEmpNo.BoundText
    txtCostCenter.Text = tdbEmpName.Columns("costcenter").Text
    txtDivision.Text = tdbEmpName.Columns("division").Text
    txtBranch.Text = tdbEmpName.Columns("branch").Text
  End If
End Sub

Private Sub Compute_hours(ByRef objTin As Object, ByRef objTout As Object, ByRef objThrs As Object)

  Dim mTime         As Double
  
  If IsDate(objTin.Text) And IsDate(objTout.Text) Then
    If CDate(objTin.Text) > CDate(objTout.Text) Then
      mTime = 24 + Format(Round(DateDiff("N", objTin.Text, "12:00 am") / 60, 2), "#,##0.00")
      objThrs.Value = Format(Round(DateDiff("N", "12:00 am", objTout.Text) / 60, 2), "#,##0.00") + mTime
    Else
      objThrs.Value = Format(Round(DateDiff("N", objTin.Text, objTout.Text) / 60, 2), "#,##0.00")
      If objThrs.Value = 0 Then objThrs.Value = 24
    End If
  Else
    objThrs.Value = 0
  End If


End Sub

Private Sub txtOTStart_LostFocus()
  Compute_hours txtOTStart, txtOTEnd, txtOTHrs
End Sub

Private Sub txtOTend_LostFocus()
  Compute_hours txtOTStart, txtOTEnd, txtOTHrs
End Sub

Private Sub txtWorkdate_LostFocus()
  
  If Trim(tdbEmpName.Text) <> "" And Not IsNull(tdbEmpName.SelectedItem) And tdbEmpNo.ApproxCount > 0 Then
    Get_Sched tdbEmpName.BoundText
  End If
    
End Sub

Private Sub Get_Sched(mEmpNo As String)
 
   Dim rsTmp          As ADODB.Recordset
  
  txtDayStat.Text = ""
  
  If IsDate(txtWorkdate.Text) Then
    NetOpen rsTmp, "", "select concat(st1in,'  ',st1out, '    ',st2in, '  ',st2out) shiftsched from dtremp " & _
                      "where shiftcode <> '' and empno = '" & mEmpNo & "' and workdate = '" & Format(txtWorkdate.Text, "YYYY-MM-DD") & "'"
                      
    If rsTmp.RecordCount > 0 Then
      txtSchedule.Text = rsTmp!shiftsched
    Else
      NetOpen rsTmp, "", "select concat(x2.t1in,'  ',x2.t1out,'    ',x2.t2in,'  ',x2.t2out) shiftsched from empshift x1 " & _
                      "left outer join shift x2 on x1.shiftcode = x2.shiftcode " & _
                      "where x1.shiftcode <> '' and empno = '" & mEmpNo & "' and dayno = '" & Weekday(txtWorkdate.Text) & "'"
                      
      If rsTmp.RecordCount > 0 Then
        txtSchedule.Text = rsTmp!shiftsched
      Else
        txtSchedule.Text = ""
        txtDayStat.Text = "Day Off"
      End If
    End If
    
    NetOpen rsTmp, "", "select * from holiday where holidaydate = '" & Format(txtWorkdate.Text, "YYYY-MM-DD") & "'"
    
    If rsTmp.RecordCount > 0 Then
      If Trim(txtDayStat.Text) = "" Then
        txtDayStat.Text = rsTmp!Holiday
      Else
        txtDayStat.Text = txtDayStat.Text & "," & rsTmp!Holiday
      End If
    End If
    
  End If
End Sub


