VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Begin VB.Form frmRptActHrsWrk 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   9240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraParmtr 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   5700
      Width           =   9615
      Begin VB.Frame fraEnd 
         BackColor       =   &H00F6F8F8&
         Caption         =   "End Date"
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
         Height          =   690
         Left            =   4230
         TabIndex        =   5
         Top             =   765
         Width           =   2595
         Begin TDBTime6Ctl.TDBTime t2 
            Height          =   300
            Left            =   1410
            TabIndex        =   6
            Top             =   255
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   529
            Caption         =   "frmRptActHrsWrk.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmRptActHrsWrk.frx":006C
            Spin            =   "frmRptActHrsWrk.frx":00BC
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
            Format          =   "hh:nn AM/PM"
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
            Text            =   "04:55 AM"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.205081018518519
         End
         Begin TDBDate6Ctl.TDBDate d2 
            Height          =   300
            Left            =   405
            TabIndex        =   7
            Top             =   255
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   529
            Calendar        =   "frmRptActHrsWrk.frx":00E4
            Caption         =   "frmRptActHrsWrk.frx":01FC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptActHrsWrk.frx":0268
            Keys            =   "frmRptActHrsWrk.frx":0286
            Spin            =   "frmRptActHrsWrk.frx":02E4
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
            Text            =   "03/13/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39520
            CenturyMode     =   0
         End
      End
      Begin VB.Frame fraStart 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Start Date"
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
         Height          =   690
         Left            =   4230
         TabIndex        =   2
         Top             =   75
         Width           =   2580
         Begin TDBTime6Ctl.TDBTime t1 
            Height          =   300
            Left            =   1395
            TabIndex        =   3
            Top             =   255
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   529
            Caption         =   "frmRptActHrsWrk.frx":030C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmRptActHrsWrk.frx":0378
            Spin            =   "frmRptActHrsWrk.frx":03C8
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
            Format          =   "hh:nn AM/PM"
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
            Text            =   "04:55 AM"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.205081018518519
         End
         Begin TDBDate6Ctl.TDBDate d1 
            Height          =   300
            Left            =   405
            TabIndex        =   4
            Top             =   255
            Width           =   960
            _Version        =   65536
            _ExtentX        =   1693
            _ExtentY        =   529
            Calendar        =   "frmRptActHrsWrk.frx":03F0
            Caption         =   "frmRptActHrsWrk.frx":0508
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptActHrsWrk.frx":0574
            Keys            =   "frmRptActHrsWrk.frx":0592
            Spin            =   "frmRptActHrsWrk.frx":05F0
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
            Text            =   "03/13/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39520
            CenturyMode     =   0
         End
      End
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   345
         Left            =   45
         TabIndex        =   8
         Tag             =   "Municipal"
         Top             =   1005
         Width           =   4050
         _ExtentX        =   7144
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
         _PropDict       =   $"frmRptActHrsWrk.frx":0618
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
      Begin OsenXPCntrl.OsenXPButton cmdView 
         Height          =   435
         Left            =   6975
         TabIndex        =   9
         Top             =   435
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   767
         BTYPE           =   5
         TX              =   "&View"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16185592
         BCOLO           =   16185592
         FCOL            =   3186872
         FCOLO           =   3186872
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmRptActHrsWrk.frx":06C2
         PICN            =   "frmRptActHrsWrk.frx":06DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBList80.TDBCombo tdbSection 
         Height          =   345
         Left            =   45
         TabIndex        =   10
         Tag             =   "Municipal"
         Top             =   300
         Width           =   4050
         _ExtentX        =   7144
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
         _PropDict       =   $"frmRptActHrsWrk.frx":0C78
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   105
         Left            =   6975
         TabIndex        =   11
         Top             =   885
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   105
         Left            =   6975
         TabIndex        =   12
         Top             =   1005
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Left            =   30
         TabIndex        =   14
         Top             =   45
         Width           =   1455
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
         Left            =   45
         TabIndex        =   13
         Top             =   750
         Width           =   1830
      End
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer xrpt 
      Height          =   5910
      Left            =   615
      TabIndex        =   15
      Top             =   300
      Width           =   8235
      lastProp        =   600
      _cx             =   14526
      _cy             =   10425
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmRptActHrsWrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CrxRep            As CRAXDRT.Report


Private Sub cmdView_Click()

    Dim CrxApp            As CRAXDRT.Application
    Dim crxDatabase       As CRAXDRT.Database
    Dim crxDatabaseTables As CRAXDRT.DatabaseTables
    Dim crxDatabaseTable  As CRAXDRT.DatabaseTable
    
    Dim rsTitoTmp   As ADODB.Recordset
    Dim rsTito      As ADODB.Recordset
    Dim rsEmployee  As ADODB.Recordset
    
    Dim mDateTmp    As String
    Dim mDate1      As String
    Dim mDate2      As String
'    Dim mEmpno      As String
    
    Dim isIn        As Boolean
    
    
    If Not IsDate(d1.Text) Then
        MsgBox "Invalid date format.", vbExclamation + vbOKOnly
        d1.SetFocus
        d1.SelStart = 0
        d1.SelLength = Len(d1.Text)
        Exit Sub
    End If
    
    If Not IsDate(t1.Text) Then
        MsgBox "Invaid time format.", vbExclamation + vbOKOnly
        t1.SetFocus
        t1.SelStart = 0
        t1.SelLength = Len(t1.Text)
        Exit Sub
    End If
    
    If Not IsDate(d2.Text) Then
        MsgBox "Invalid date format.", vbExclamation + vbOKOnly
        d2.SetFocus
        d2.SelStart = 0
        d2.SelLength = Len(d2.Text)
        Exit Sub
    End If
    
    If Not IsDate(t2.Text) Then
        MsgBox "Invalid time format.", vbExclamation + vbOKOnly
        t2.SetFocus
        t2.SelStart = 0
        t2.SelLength = Len(t2.Text)
        Exit Sub
    End If
        
    mDate1 = Format(d1.Text & " " & t1.Text, "YYYY-MM-DD hh:nn:ss")
    mDate2 = Format(d2.Text & " " & t2.Text, "YYYY-MM-DD hh:nn:ss")
    
    If CDate(mDate1) > CDate(mDate2) Then
        MsgBox "End date must be greater than start date.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    
    If Not IsNull(tdbSection.SelectedItem) And Trim(tdbSection.Text) <> "" And tdbSection.ApproxCount > 0 Then
        If IsNull(tdbEmployee.SelectedItem) Or Trim(tdbEmployee.Text) = "" Or tdbEmployee.ApproxCount = 0 Then
            NetOpen rsEmployee, "select employeecode from employee where sectioncode = '" & tdbSection.BoundText & "'"
        Else
            NetOpen rsEmployee, "select employeecode from employee where employeecode = '" & tdbEmployee.BoundText & "'"
        End If
    Else
        NetOpen rsEmployee, "select employeecode from employee"
    End If
    

    
    If rsEmployee.RecordCount > 0 Then
    
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        
        ConMain.Execute "delete from rptacthrswrk"
    
        rsEmployee.MoveFirst
        
        pb1.Max = rsEmployee.RecordCount
        pb1.Value = 0
                
        Do While Not rsEmployee.EOF
        
            Set rsTitoTmp = Nothing
            Set rsTitoTmp = New ADODB.Recordset
            
            With rsTitoTmp
                .Fields.Append "wrkdate", adDate
                .Fields.Append "tin", adDate
                .Fields.Append "tout", adDate
                .Fields.Append "ttlhrswrk", adDouble
                .Open
                .Sort = "wrkdate"
            End With
        
            pb1.Value = pb1.Value + 1
            NetOpen rsTito, "select * from tito where employeecode = '" & rsEmployee!employeecode & "' and complog between " & _
                                "'" & mDate1 & "' and '" & mDate2 & "' order by complog"
        
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
                                rsTitoTmp.Fields("ttlhrswrk") = DiffHrs(rsTitoTmp.Fields("tin"), rsTitoTmp.Fields("tout"))
                                rsTitoTmp.Update
                                rsTitoTmp.AddNew
                                rsTitoTmp.Fields("wrkdate") = !datelog
                                rsTitoTmp.Fields("tin") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                                isIn = True
                            End If
                            
                        Else
                        
                            If isIn Then
                                rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                                rsTitoTmp.Fields("ttlhrswrk") = DiffHrs(rsTitoTmp.Fields("tin"), rsTitoTmp.Fields("tout"))
                                rsTitoTmp.Update
                                isIn = False
                                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                            Else
                                rsTitoTmp.AddNew
                                rsTitoTmp.Fields("wrkdate") = Format(mDateTmp, "MM/DD/YYYY")
                                rsTitoTmp.Fields("tin") = mDateTmp
                                rsTitoTmp.Fields("tout") = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                                rsTitoTmp.Fields("ttlhrswrk") = DiffHrs(rsTitoTmp.Fields("tin"), rsTitoTmp.Fields("tout"))
                                rsTitoTmp.Update
                                isIn = False
                                mDateTmp = Format(CDate(!datelog) & " " & !timelog, "mm/dd/yy hh:nn:ss")
                            End If
                            
                        End If
                        .MoveNext
                        DoEvents
                    Loop

                    rsTitoTmp.MoveLast
                    If rsTitoTmp!tout = "" Then
                        rsTitoTmp.Delete
                        rsTitoTmp.Update
                    End If
                    
        
                    If rsTitoTmp.RecordCount > 0 Then
                        rsTitoTmp.MoveFirst
                        pb2.Max = rsTitoTmp.RecordCount
                        pb2.Value = 0
                        Do While Not rsTitoTmp.EOF
                            pb2.Value = pb2.Value + 1
                            ConMain.Execute "insert into rptacthrswrk(employeecode,workdate,timein,timeout,ttlhrs) values " & _
                                                "('" & rsEmployee!employeecode & "','" & Format(rsTitoTmp!wrkdate, "YYYY-MM-DD") & "', " & _
                                                "'" & Format(rsTitoTmp!tin, "YYYY-MM-DD hh:nn:ss") & "','" & Format(rsTitoTmp!tout, "YYYY-MM-DD hh:nn:ss") & "', " & _
                                                IIf(Not IsNumeric(rsTitoTmp!ttlhrswrk), 0, rsTitoTmp!ttlhrswrk) & ")"
                            rsTitoTmp.MoveNext
                        Loop
                    End If
                End If
        
            End With
            rsEmployee.MoveNext
            DoEvents
        Loop
        
        ConMain.CommitTrans
        
        Set CrxApp = New CRAXDRT.Application
        Set CrxRep = New CRAXDRT.Report
        
        Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\ActHrsWrk.rpt")
        
            Set crxDatabase = CrxRep.Database
            Set crxDatabaseTables = crxDatabase.Tables

            For Each crxDatabaseTable In crxDatabaseTables
                crxDatabaseTable.ConnectionProperties("data source name").Value = SQLDatabase
                crxDatabaseTable.ConnectionProperties("user id").Value = SQLUsername
                crxDatabaseTable.ConnectionProperties("password").Value = SQLPassword
            Next crxDatabaseTable
        
        CrxRep.ParameterFields.GetItemByName("mreporttitle").AddCurrentValue "Actual Time Log Report"
        CrxRep.ParameterFields.GetItemByName("startdate").AddCurrentValue CDate(mDate1)
        CrxRep.ParameterFields.GetItemByName("enddate").AddCurrentValue CDate(mDate2)
        xrpt.ReportSource = CrxRep
        xrpt.ViewReport
        xrpt.Zoom 130
        
        Set crxDatabase = Nothing
        Set crxDatabaseTable = Nothing
        Set crxDatabaseTables = Nothing
        Set CrxApp = Nothing
        
        pb1.Value = 0
        pb2.Value = 0
        
    Else
        MsgBox "No record found.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub d1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub d1_GotFocus()
    d1.SelStart = 0
    d1.SelLength = Len(d1.Text)
End Sub

Private Sub d2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub d2_GotFocus()
    d2.SelStart = 0
    d2.SelLength = Len(d2.Text)
End Sub

Private Sub Form_Load()

    Add_MDIButton Me.Name, TitleBar.Caption
    
    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    SendMessage pb2.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb2.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    bind_tdb ConMain, tdbSection, "select sectioncode,sectionname from section order by sectionname", "sectionname", "sectioncode"
    
    d1.Text = Format(Now, "mm/dd/yyyy")
    d2.Text = Format(Now, "mm/dd/yyyy")
    t1.Text = "12:00 AM"
    t2.Text = "12:00 AM"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    TitleBar.Move 0, 0, Me.ScaleWidth

    With xrpt
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - (.Top + fraParmtr.Height)
    End With
    
    With fraParmtr
        .Top = xrpt.Top + xrpt.Height
        .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    End With
    
End Sub

Private Sub t1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub t1_GotFocus()
    t1.SelStart = 0
    t1.SelLength = Len(t1.Text)
End Sub

Private Sub t2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub t2_GotFocus()
    t2.SelStart = 0
    t2.SelLength = Len(t2.Text)
End Sub

Private Sub tdbEmployee_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbEmployee, tdbEmployee.RowSource, tdbEmployee.Text
    End If
    
End Sub

Private Function DiffHrs(mHrs1 As Date, mHrs2 As Date) As Double
    DiffHrs = Format(Round(DateDiff("N", mHrs1, mHrs2) / 60, 2), "#,##0.00")
End Function


Private Sub tdbSection_ItemChange()
    bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname, ', ', firstname,' ',middlename) fullname from employee " & _
        "where sectioncode = '" & tdbSection.BoundText & "' order by concat(lastname, ', ', firstname,' ',middlename)", "fullname", "employeecode"
    If tdbEmployee.ApproxCount = 0 Then
        tdbEmployee.BoundText = ""
    End If
End Sub

Private Sub tdbsection_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbSection, tdbSection.RowSource, tdbSection.Text
        tdbSection_ItemChange
    End If
End Sub

Private Sub xrpt_PrintButtonClicked(UseDefault As Boolean)
    CrxRep.PrinterSetup Me.hwnd
End Sub
