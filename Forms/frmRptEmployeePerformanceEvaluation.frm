VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Begin VB.Form frmRptEmployeePerformanceEvaluation 
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   9945
   Tag             =   "Employee Performance Evaluation Report"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraParmtr 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   90
      TabIndex        =   1
      Top             =   1005
      Width           =   3465
      Begin VB.OptionButton optSummary 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   2025
         Width           =   1305
      End
      Begin VB.Frame fraSummary 
         BackColor       =   &H00F6F8F8&
         Height          =   2340
         Left            =   60
         TabIndex        =   10
         Top             =   2190
         Width           =   3360
         Begin VB.OptionButton optEvaluationPeriod 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            Caption         =   "Period of Evaluation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   60
            TabIndex        =   19
            Top             =   1500
            Width           =   3015
         End
         Begin VB.OptionButton optEvaluationDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            Caption         =   "Date of Evaluation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   60
            TabIndex        =   18
            Top             =   810
            Width           =   2115
         End
         Begin VB.OptionButton optInputDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00F6F8F8&
            Caption         =   "Input Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   60
            TabIndex        =   17
            Top             =   225
            Value           =   -1  'True
            Width           =   1350
         End
         Begin TDBDate6Ctl.TDBDate txtEvalDate 
            Height          =   315
            Left            =   345
            TabIndex        =   12
            Top             =   1140
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            Calendar        =   "frmRptEmployeePerformanceEvaluation.frx":0000
            Caption         =   "frmRptEmployeePerformanceEvaluation.frx":0106
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptEmployeePerformanceEvaluation.frx":016C
            Keys            =   "frmRptEmployeePerformanceEvaluation.frx":018A
            Spin            =   "frmRptEmployeePerformanceEvaluation.frx":01E8
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
            Text            =   "09/29/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39720
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate txtEvalFrom 
            Height          =   315
            Left            =   345
            TabIndex        =   13
            Top             =   1830
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            Calendar        =   "frmRptEmployeePerformanceEvaluation.frx":0210
            Caption         =   "frmRptEmployeePerformanceEvaluation.frx":0316
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptEmployeePerformanceEvaluation.frx":037C
            Keys            =   "frmRptEmployeePerformanceEvaluation.frx":039A
            Spin            =   "frmRptEmployeePerformanceEvaluation.frx":03F8
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
            Text            =   "09/29/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39720
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate txtEvalTo 
            Height          =   315
            Left            =   1860
            TabIndex        =   14
            Top             =   1830
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            Calendar        =   "frmRptEmployeePerformanceEvaluation.frx":0420
            Caption         =   "frmRptEmployeePerformanceEvaluation.frx":0526
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptEmployeePerformanceEvaluation.frx":058C
            Keys            =   "frmRptEmployeePerformanceEvaluation.frx":05AA
            Spin            =   "frmRptEmployeePerformanceEvaluation.frx":0608
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
            Text            =   "09/29/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39720
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate txtTrnxdate 
            Height          =   315
            Left            =   360
            TabIndex        =   15
            Top             =   465
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            Calendar        =   "frmRptEmployeePerformanceEvaluation.frx":0630
            Caption         =   "frmRptEmployeePerformanceEvaluation.frx":0736
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRptEmployeePerformanceEvaluation.frx":079C
            Keys            =   "frmRptEmployeePerformanceEvaluation.frx":07BA
            Spin            =   "frmRptEmployeePerformanceEvaluation.frx":0818
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
            Text            =   "09/29/2008"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   39720
            CenturyMode     =   0
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   1500
            TabIndex        =   16
            Top             =   1875
            Width           =   345
         End
      End
      Begin VB.OptionButton optIndividual 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   45
         TabIndex        =   9
         Top             =   150
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.Frame fraIndividual 
         BackColor       =   &H00F6F8F8&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1545
         Left            =   45
         TabIndex        =   4
         Top             =   300
         Width           =   3390
         Begin TrueOleDBList80.TDBCombo tdbEmployee 
            Height          =   345
            Left            =   75
            TabIndex        =   5
            Tag             =   "Municipal"
            Top             =   345
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   609
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            _EDITHEIGHT     =   609
            _GAPHEIGHT      =   53
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "costcentercode"
            Columns(0).DataField=   "costcentercode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "desciption"
            Columns(1).DataField=   "costcenter"
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
            _PropDict       =   $"frmRptEmployeePerformanceEvaluation.frx":0840
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
         Begin TrueOleDBList80.TDBCombo tdbEmpEval 
            Height          =   345
            Left            =   75
            TabIndex        =   6
            Tag             =   "Municipal"
            Top             =   1080
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   609
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            _EDITHEIGHT     =   609
            _GAPHEIGHT      =   53
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "costcentercode"
            Columns(0).DataField=   "costcentercode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "desciption"
            Columns(1).DataField=   "costcenter"
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
            _PropDict       =   $"frmRptEmployeePerformanceEvaluation.frx":08EA
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
         Begin VB.Label Label1 
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
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   8
            Top             =   135
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of evaluation"
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
            Left            =   75
            TabIndex        =   7
            Top             =   870
            Width           =   2415
         End
      End
      Begin OsenXPCntrl.OsenXPButton cmdView 
         Height          =   435
         Left            =   435
         TabIndex        =   0
         Top             =   4890
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
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmRptEmployeePerformanceEvaluation.frx":0994
         PICN            =   "frmRptEmployeePerformanceEvaluation.frx":09B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer xrpt 
      Height          =   8070
      Left            =   3870
      TabIndex        =   2
      Top             =   -30
      Width           =   6405
      lastProp        =   600
      _cx             =   11298
      _cy             =   14235
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
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   609
      Caption         =   "Employee Performance Evaluation Report"
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
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "frmRptEmployeePerformanceEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CrxRep                  As CRAXDRT.Report

Private Sub cmdView_Click()

    Dim CrxApp              As CRAXDRT.Application
    Dim crxDatabase         As CRAXDRT.Database
    Dim crxDatabaseTable    As CRAXDRT.DatabaseTable
    Dim crxDatabaseTables   As CRAXDRT.DatabaseTables
    
    Dim mAddQuery           As String
    
    
    If optIndividual.Value = True Then
    
      If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
        mAddQuery = " where a.employeecode = " & tdbEmployee.BoundText & ""
      Else
        MsgBox "Please select an employee.", vbExclamation + vbOKOnly
        tdbEmployee.SetFocus
        Exit Sub
      End If
      If Trim(tdbEmpEval.Text) <> "" And Not IsNull(tdbEmpEval.SelectedItem) And tdbEmpEval.ApproxCount > 0 Then
        mAddQuery = mAddQuery & " and a.empevalcode = " & tdbEmpEval.BoundText & " "
      Else
        MsgBox "Please select an evaluation date.", vbExclamation + vbOKOnly
        tdbEmpEval.SetFocus
        Exit Sub
      End If
      
    Else
      
      If optInputDate.Value = True Then
        mAddQuery = " where a.trnxdate = '" & Format(txtTrnxdate.Text, "YYYY-MM-DD") & "'"
      ElseIf optEvaluationDate.Value = True Then
        mAddQuery = " where a.evaluationdate = '" & Format(txtEvalDate.Text, "YYYY-MM-DD") & "'"
      Else
        mAddQuery = " where a.evalfrom >= '" & Format(txtEvalFrom.Text, "YYYY-MM-DD") & "' and a.evalto <= '" & Format(txtEvalTo.Text, "YYYY-MM-DD") & "'"
      End If
      
    End If
    
    
    Set CrxApp = Nothing
    Set CrxRep = Nothing
    
    Set CrxApp = New CRAXDRT.Application
    Set CrxRep = New CRAXDRT.Report
    
    If optIndividual.Value = True Then
      Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\EmployeePerfomanceEvaluation.rpt")
    Else
      Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\EmployeePerfomanceEvaluationSummary.rpt")
    End If
    Set crxDatabase = CrxRep.Database
    Set crxDatabaseTables = crxDatabase.Tables
    
    For Each crxDatabaseTable In crxDatabaseTables
        crxDatabaseTable.ConnectionProperties("connection string") = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & SQLServerName & "; DATABASE=" & SQLDatabase & "; UID=" & SQLUsername & "; PWD=" & SQLPassword & "; PORT='" & SQLPort & "'"
    Next crxDatabaseTable
    
    CrxRep.ParameterFields.GetItemByName("mAddQuery").AddCurrentValue mAddQuery
    
    xrpt.ReportSource = CrxRep
    xrpt.ViewReport
    xrpt.Zoom 100
    
    Set crxDatabase = Nothing
    Set crxDatabaseTable = Nothing
    Set crxDatabaseTables = Nothing
    Set CrxApp = Nothing
    
End Sub

Private Sub Form_Load()

    Add_MDIButton Me.Name, TitleBar.Caption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fraParmtr
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Height = Me.ScaleHeight - .Top
    End With

    With xrpt
        .Top = TitleBar.Top + TitleBar.Height
        .Left = fraParmtr.Left + fraParmtr.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top
    End With
   
    fraIndividual.Enabled = True
    fraSummary.Enabled = False
    txtTrnxdate.Text = Format(Now, "MM/DD/YYYY")
    txtEvalDate.Text = Format(Now, "MM/DD/YYYY")
    txtEvalFrom.Text = Format(Now, "MM/DD/YYYY")
    txtEvalTo.Text = Format(Now, "MM/DD/YYYY")
End Sub

Private Sub optIndividual_Click()
  If optIndividual.Value = True Then
    fraIndividual.Enabled = True
    fraSummary.Enabled = False
  End If
End Sub

Private Sub optSummary_Click()
  If optSummary.Value = True Then
    fraIndividual.Enabled = False
    fraSummary.Enabled = True
  End If
End Sub

Private Sub tdbEmpEval_GotFocus()

  If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
    
    If Trim(tdbEmpEval.Text) <> "" And Not IsNull(tdbEmpEval.SelectedItem) And tdbEmpEval.ApproxCount > 0 Then
        tdbEmpEval.Tag = tdbEmpEval.BoundText
    Else
        tdbEmpEval.Tag = ""
    End If
    
    bind_tdb ConMain, tdbEmpEval, "select empevalcode,evaluationdate from empeval where employeecode = " & tdbEmployee.BoundText & " order by evaluationdate", "evaluationdate", "empevalcode"
    
    tdbEmpEval.BoundText = tdbEmpEval.Tag
    
    If IsNull(tdbEmpEval.SelectedItem) Then tdbEmpEval.BoundText = ""
    
    tdbEmpEval.Tag = ""
  
  Else
    
    bind_tdb ConMain, tdbEmpEval, "select empevalcode,evaluationdate from empeval where empevalcode = 0 limit 0", "evaluationdate", "empevalcode"
    
  End If

End Sub

Private Sub tdbEmpEval_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEmpEval, tdbEmpEval.RowSource, tdbEmpEval.Text
  End If
End Sub

Private Sub tdbEmployee_GotFocus()
    
    If Trim(tdbEmployee.Text) <> "" And Not IsNull(tdbEmployee.SelectedItem) And tdbEmployee.ApproxCount > 0 Then
        tdbEmployee.Tag = tdbEmployee.BoundText
    Else
        tdbEmployee.Tag = ""
    End If
    
    bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname,', ',firstname,' ',middlename) employeename from employee order by employeename", "employeename", "employeecode"
    
    tdbEmployee.BoundText = tdbEmployee.Tag

End Sub

Private Sub tdbEmployee_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEmployee, tdbEmployee.RowSource, tdbEmployee.Text
  End If
End Sub

Private Sub tdbEmployee_LostFocus()
  
    If tdbEmployee.Tag <> tdbEmployee.BoundText Then tdbEmpEval.BoundText = ""
    
End Sub

Private Sub txtEvalDate_GotFocus()
  With txtEvalDate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEvalFrom_GotFocus()
  With txtEvalFrom
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEvalTo_GotFocus()
  With txtEvalTo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtTrnxdate_GotFocus()
  With txtTrnxdate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub xrpt_PrintButtonClicked(UseDefault As Boolean)
    CrxRep.PrinterSetup Me.hwnd
End Sub
