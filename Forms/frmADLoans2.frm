VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADLoans2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Loan"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
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
   Icon            =   "frmADLoans2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   3450
      Left            =   30
      TabIndex        =   13
      Top             =   -75
      Width           =   7890
      Begin TDBText6Ctl.TDBText txtLoanno 
         Height          =   300
         Left            =   2040
         TabIndex        =   0
         Top             =   195
         Width           =   2745
         _Version        =   65536
         _ExtentX        =   4842
         _ExtentY        =   529
         Caption         =   "frmADLoans2.frx":6852
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":68BE
         Key             =   "frmADLoans2.frx":68DC
         BackColor       =   14737632
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
      Begin TDBNumber6Ctl.TDBNumber txtLoanAmnt 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   1980
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmADLoans2.frx":6920
         Caption         =   "frmADLoans2.frx":6940
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":69AC
         Keys            =   "frmADLoans2.frx":69CA
         Spin            =   "frmADLoans2.frx":6A14
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
         ForeColor       =   4210752
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBNumber6Ctl.TDBNumber txtDedPerPayDay 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   2340
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmADLoans2.frx":6A3C
         Caption         =   "frmADLoans2.frx":6A5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":6AC8
         Keys            =   "frmADLoans2.frx":6AE6
         Spin            =   "frmADLoans2.frx":6B30
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
         ForeColor       =   4210752
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBNumber6Ctl.TDBNumber txtNoofInst 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   2700
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmADLoans2.frx":6B58
         Caption         =   "frmADLoans2.frx":6B78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":6BDE
         Keys            =   "frmADLoans2.frx":6BFC
         Spin            =   "frmADLoans2.frx":6C46
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   4210752
         Format          =   "#,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1961820165
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TDBDate6Ctl.TDBDate txtLoandate 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   1620
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   556
         Calendar        =   "frmADLoans2.frx":6C6E
         Caption         =   "frmADLoans2.frx":6D74
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":6DDA
         Keys            =   "frmADLoans2.frx":6DF8
         Spin            =   "frmADLoans2.frx":6E56
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
         ForeColor       =   4210752
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
         Text            =   "01/16/2008"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39463
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate txtStartdate 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   3060
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calendar        =   "frmADLoans2.frx":6E7E
         Caption         =   "frmADLoans2.frx":6F84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":6FEA
         Keys            =   "frmADLoans2.frx":7008
         Spin            =   "frmADLoans2.frx":7066
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
         ForeColor       =   4210752
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
         Text            =   "01/16/2008"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39463
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText txtRemarks 
         Height          =   1755
         Left            =   4755
         TabIndex        =   9
         Tag             =   "txtRegistrationRemarks"
         Top             =   1620
         Width           =   3060
         _Version        =   65536
         _ExtentX        =   5397
         _ExtentY        =   3096
         Caption         =   "frmADLoans2.frx":708E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":70FA
         Key             =   "frmADLoans2.frx":7118
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   4210752
         ReadOnly        =   0
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   -1
         ScrollBars      =   2
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   200
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin TrueOleDBList80.TDBCombo tdbLoanTypes 
         Height          =   345
         Left            =   2040
         TabIndex        =   3
         Tag             =   "Municipal"
         Top             =   1230
         Width           =   2985
         _ExtentX        =   5265
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
         _PropDict       =   $"frmADLoans2.frx":715C
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
      Begin TDBText6Ctl.TDBText txtReferenceNo 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   885
         Width           =   2745
         _Version        =   65536
         _ExtentX        =   4842
         _ExtentY        =   529
         Caption         =   "frmADLoans2.frx":7206
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":7272
         Key             =   "frmADLoans2.frx":7290
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   4210752
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
         MaxLength       =   20
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
      Begin TDBText6Ctl.TDBText txtFullname 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   540
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   529
         Caption         =   "frmADLoans2.frx":72D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADLoans2.frx":7340
         Key             =   "frmADLoans2.frx":735E
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   4210752
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
         MaxLength       =   20
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
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
         Left            =   450
         TabIndex        =   23
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   465
         TabIndex        =   22
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   240
         Left            =   435
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction per Payday"
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
         Height          =   240
         Left            =   -105
         TabIndex        =   20
         Top             =   2385
         Width           =   2025
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Application Date"
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
         Left            =   450
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Type"
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
         Left            =   465
         TabIndex        =   18
         Top             =   1275
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan No."
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
         Height          =   240
         Left            =   435
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   -105
         TabIndex        =   16
         Top             =   3120
         Width           =   2025
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Installments"
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
         Height          =   240
         Index           =   0
         Left            =   -105
         TabIndex        =   15
         Top             =   2775
         Width           =   2025
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   2385
         Width           =   765
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -15
      TabIndex        =   12
      Top             =   3405
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   210
         Left            =   4110
         TabIndex        =   24
         Top             =   165
         Visible         =   0   'False
         Width           =   990
      End
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   60
         TabIndex        =   11
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
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
         Image           =   "frmADLoans2.frx":73A2
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   1800
         TabIndex        =   10
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         Caption         =   "&OK"
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
         Image           =   "frmADLoans2.frx":807C
         cBack           =   14737632
      End
   End
End
Attribute VB_Name = "frmADLoans2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mEmployeeCode        As Integer

Public mNew                 As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsDateTime          As ADODB.Recordset

    Dim mLoanDedCode        As Integer

    Dim mLoanCode           As Integer
    
    If Trim(tdbLoanTypes.Text) = "" Or IsNull(tdbLoanTypes.SelectedItem) Or tdbLoanTypes.ApproxCount = 0 Then
      MsgBox "Please select a loan type.", vbExclamation + vbOKOnly
      tdbLoanTypes.SetFocus
      Exit Sub
    End If

    If Not IsDate(txtLoandate.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtLoandate.SetFocus
      Exit Sub
    End If

    If Not IsNumeric(txtLoanAmnt.Text) Then
      MsgBox "Please enter a vaild number.", vbExclamation + vbOKOnly
      txtDedPerPayDay.SetFocus
      Exit Sub
    Else
      If CDbl(txtLoanAmnt.Text) <= 0 Then
        MsgBox "Please enter an amount greater than zero.", vbExclamation + vbOKOnly
        txtLoanAmnt.SetFocus
        Exit Sub
      End If
    End If

    If Not IsNumeric(txtDedPerPayDay.Text) Then
      MsgBox "Please enter a vaild number.", vbExclamation + vbOKOnly
      txtDedPerPayDay.SetFocus
      Exit Sub
    Else
      If CDbl(txtDedPerPayDay.Text) <= 0 Then
        MsgBox "Please enter an amount greater than zero.", vbExclamation + vbOKOnly
        txtDedPerPayDay.SetFocus
        Exit Sub
      End If
    End If

    If Not IsDate(txtStartDate.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtStartDate.SetFocus
      Exit Sub
    End If

    If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If

    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"

    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans

    If mNew = True Then

        txtLoanno.Text = LastCode("Loans")
        
        mLoanCode = txtLoanno.Text

        ConMain.Execute "insert into loans(loancode,dummycode,employeecode,loantypescode,costcentercode," & _
                            "divisioncode,branchcode,loandate,loanamnt,dedperpayday, " & _
                            "noofinst,startdate,status,remarks,referenceno) values (" & _
                            txtLoanno.Text & ",'" & Format(txtLoanno.Text, "0000000000") & "', " & mEmployeeCode & "," & tdbLoanTypes.BoundText & "," & frmADLoans.mCostCenterCode & ", " & _
                            frmADLoans.mDivisionCode & "," & frmADLoans.mBranchCode & ",'" & Format(txtLoandate.Text, "YYYY-MM-DD") & "'," & Format(txtLoanAmnt.Text, "##0.00") & "," & Format(txtDedPerPayDay.Text, "##0.00") & ", " & _
                            txtNoofInst.Text & ",'" & Format(txtStartDate.Text, "YYYY-MM-DD") & "'," & "'Active','" & Swap(txtRemarks.Text) & "','" & Swap(txtReferenceNo.Text) & "')"

        'mLoanDedCode = LastCode("LoanDed")
        
        mLoanDedCode = LastLoanCodeUsed(mLoanCode)
        ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled,usercode) values " & _
                      "(" & mLoanDedCode & "," & txtLoanno.Text & "," & tdbLoanTypes.BoundText & "," & mEmployeeCode & "," & _
                       0 & ",'" & Format(rsDateTime!currentdate, "YYYY-MM-DD") & "', 0 ," & Format(txtLoanAmnt.Text, "##0.00") & ",'Y','N'," & GlobalUserID & ")"


    Else
        
        mLoanCode = frmADLoans.rsLoans!loancode
        
        ConMain.Execute "update loans set loantypescode = " & tdbLoanTypes.BoundText & ", costcentercode = " & frmADLoans.mCostCenterCode & ", " & _
                            "divisioncode = " & frmADLoans.mDivisionCode & ",branchcode = " & frmADLoans.mBranchCode & ", " & _
                            "loandate = '" & Format(txtLoandate.Text, "YYYY-MM-DD") & "', loanamnt = " & Format(txtLoanAmnt.Text, "##0.00") & ", dedperpayday = " & Format(txtDedPerPayDay.Text, "##0.00") & ", " & _
                            "noofinst = " & Format(txtNoofInst.Text, "##0.00") & ", startdate = '" & Format(txtStartDate.Text, "YYYY-MM-DD") & "',remarks = '" & Swap(txtRemarks.Text) & "',referenceno = '" & Swap(txtReferenceNo.Text) & "' where loancode = " & frmADLoans.rsLoans!loancode & ""

    End If

    ConMain.CommitTrans
    
    Me.MousePointer = vbHourglass
    With frmADLoans
        .rsLoans.Requery
        .Get_LoanSum
        .rsLoans.MoveFirst
        .rsLoans.Find "loancode = " & mLoanCode & ""
    End With
    Me.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Command1_Click()
  
    Command1.Enabled = False
    Dim mLoan_Ins           As ADODB.Recordset
    
    Dim mLoanDedCode        As Integer
    Dim mLoanCode           As Integer
    Dim mLoanTypeID         As Integer
    
    mLoanTypeID = 131
    
    NetOpen mLoan_Ins, "select * from loan_ins"
    
    If mLoan_Ins.RecordCount > 0 Then
    
      ConMain.Execute "set autocommit = 0"
      ConMain.BeginTrans
    
      With mLoan_Ins
        .MoveFirst
        Do While Not .EOF
        
          mLoanCode = LastCode("Loans")
          ConMain.Execute "insert into loans(loancode,dummycode,employeecode,loantypescode,costcentercode," & _
                            "divisioncode,branchcode,loandate,loanamnt,dedperpayday, " & _
                            "noofinst,startdate,status,remarks,referenceno) values (" & _
                            mLoanCode & ",'" & Format(mLoanCode, "0000000000") & "', " & .Fields("empid").Value & "," & mLoanTypeID & "," & .Fields("costcentercode").Value & ", " & _
                            .Fields("divisioncode").Value & "," & .Fields("branchcode").Value & ",'" & Format(.Fields("appdate").Value, "YYYY-MM-DD") & "'," & Format(.Fields("loanamt").Value, "##0.00") & "," & Format(.Fields("dedperpay").Value, "##0.00") & "," & _
                            "24,'" & Format(.Fields("appdate").Value, "YYYY-MM-DD") & "'," & "'Active','','" & Swap(txtReferenceNo.Text) & "')"
          
          mLoanDedCode = LastLoanCodeUsed(mLoanCode)
          ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled,usercode) values " & _
                          "(" & mLoanDedCode & "," & mLoanCode & "," & mLoanTypeID & "," & .Fields("empid").Value & "," & _
                          0 & ",'" & Format(.Fields("appdate").Value, "YYYY-MM-DD") & "', 0 ," & Format(.Fields("loanamt").Value, "##0.00") & ",'Y','N'," & GlobalUserID & ")"


          .MoveNext
        Loop
        
        MsgBox "Import Comlete"
      End With
      
      ConMain.CommitTrans
      
    End If
    Command1.Enabled = True
End Sub

Private Sub Form_Activate()
        txtReferenceNo.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsChk           As ADODB.Recordset
    
    If mNew = True Then
        txtLoandate.Text = Format(Now, "MM/DD/YYYY")
        txtStartDate.Text = Format(Now, "MM/DD/YYYY")
    Else
        NetOpen rsChk, "select * from loanded where loancode = " & frmADLoans.rsLoans!loancode & " limit 2"
        bind_tdb ConMain, tdbLoanTypes, "select loantypescode, loantypesname from loantypes order by loantypesname", "loantypesname", "loantypescode"
        If rsChk.RecordCount > 1 Then
            'txtLoandate.Enabled = False
            tdbLoanTypes.Enabled = False
            txtLoanAmnt.Enabled = False
            txtStartDate.Enabled = False
        End If
    End If
End Sub


Private Sub txtFullname_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub tdbLoanTypes_GotFocus()
    With tdbLoanTypes
        .Tag = .BoundText
        bind_tdb ConMain, tdbLoanTypes, "select loantypescode, loantypesname from loantypes order by loantypesname", "loantypesname", "loantypescode"
        .BoundText = .Tag
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tdbLoanTypes_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbLoanTypes, tdbLoanTypes.RowSource, tdbLoanTypes.Text
  End If
End Sub

Private Sub Compute_Inst()
  
    Dim mInst         As Double
    Dim i             As Integer
    
    i = 0
    If IsNumeric(txtLoanAmnt.Text) And IsNumeric(txtDedPerPayDay.Text) Then
        If CDbl(txtLoanAmnt.Text) > 0 And CDbl(txtDedPerPayDay.Text) > 0 Then
            
            mInst = CDbl(txtLoanAmnt.Text)
            
            Do While mInst > 0
                
                If mInst >= CDbl(txtDedPerPayDay.Text) Then
                    mInst = mInst - CDbl(txtDedPerPayDay.Text)
                    i = i + 1
                Else
                    Exit Do
                End If
                
            Loop
            
            If i > 1 Then
                If mInst > 0 Then
                    If (mInst / CDbl(txtDedPerPayDay.Text)) > 0.1 Then
'                        I = I - 1
'                    Else
                        i = i + 1
                    End If
                End If
            End If
            
        End If
    End If
     txtNoofInst.Text = i
End Sub

Private Sub txtDedPerPayDay_GotFocus()
    With txtDedPerPayDay
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDedPerPayDay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDedPerPayDay_LostFocus()
  Compute_Inst
End Sub

Private Sub txtLoanAmnt_GotFocus()
    With txtLoanAmnt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLoanAmnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoanAmnt_LostFocus()
  Compute_Inst
End Sub

Private Sub txtLoandate_GotFocus()
    With txtLoandate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLoandate_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoanno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNoofInst_GotFocus()
    With txtNoofInst
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNoofInst_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRemarks_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtReferenceNo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtStartDate_GotFocus()
    With txtStartDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtStartDate_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Clear_Texts()

    Dim rsDateTime          As ADODB.Recordset
    
    
    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"
    
    mNew = True
    
    txtLoanno.Text = ""
    txtLoandate.Text = Format(rsDateTime!currentdate, "MM/DD/YYYY")
    tdbLoanTypes.BoundText = ""
    txtLoanAmnt.Text = "0.00"
    txtDedPerPayDay.Text = "0.00"
    txtNoofInst.Text = "0"
    txtStartDate.Text = Format(rsDateTime!currentdate, "MM/DD/YYYY")
    txtRemarks.Text = ""
    txtReferenceNo.Text = ""
    
    
    
    
End Sub

