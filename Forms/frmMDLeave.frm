VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDLeave 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   8865
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSearch 
      BackColor       =   &H00F6F8F8&
      Height          =   600
      Left            =   135
      TabIndex        =   14
      Top             =   450
      Width           =   8040
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   1545
         TabIndex        =   15
         Top             =   195
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   529
         Caption         =   "frmMDLeave.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":006C
         Key             =   "frmMDLeave.frx":008A
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH"
         Height          =   255
         Left            =   540
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame frmeLeaveTypes 
      BackColor       =   &H00F6F8F8&
      Enabled         =   0   'False
      Height          =   2130
      Left            =   105
      TabIndex        =   6
      Top             =   885
      Width           =   7275
      Begin VB.CheckBox chkWpay 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         Caption         =   "With pay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4185
         TabIndex        =   25
         Top             =   1395
         Width           =   2010
      End
      Begin TDBNumber6Ctl.TDBNumber txtLeaveCredits 
         Height          =   300
         Left            =   1770
         TabIndex        =   3
         Top             =   1350
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         Calculator      =   "frmMDLeave.frx":00CE
         Caption         =   "frmMDLeave.frx":00EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":015A
         Keys            =   "frmMDLeave.frx":0178
         Spin            =   "frmMDLeave.frx":01C2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText txtLeaveTypesCode 
         Height          =   300
         Left            =   1770
         TabIndex        =   0
         Top             =   210
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3528
         _ExtentY        =   529
         Caption         =   "frmMDLeave.frx":01EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":0256
         Key             =   "frmMDLeave.frx":0274
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
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
         Text            =   "AUTO GENERATED..."
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
      Begin TDBText6Ctl.TDBText txtDescription 
         Height          =   300
         Left            =   1770
         TabIndex        =   2
         Top             =   960
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   529
         Caption         =   "frmMDLeave.frx":02B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":0324
         Key             =   "frmMDLeave.frx":0342
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
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
      Begin TDBText6Ctl.TDBText txtLeaveTypes 
         Height          =   300
         Left            =   1770
         TabIndex        =   1
         Top             =   585
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   529
         Caption         =   "frmMDLeave.frx":0386
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":03F2
         Key             =   "frmMDLeave.frx":0410
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
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
      Begin TDBNumber6Ctl.TDBNumber txtConvertibleDays 
         Height          =   300
         Left            =   1770
         TabIndex        =   4
         Top             =   1725
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         Calculator      =   "frmMDLeave.frx":0454
         Caption         =   "frmMDLeave.frx":0474
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMDLeave.frx":04E0
         Keys            =   "frmMDLeave.frx":04FE
         Spin            =   "frmMDLeave.frx":0548
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "day(s)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3285
         TabIndex        =   13
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "day(s)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3285
         TabIndex        =   12
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Conversion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   1710
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Credit(s)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   1365
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   975
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   615
         Width           =   1470
      End
   End
   Begin TrueOleDBGrid80.TDBGrid gridLeaveTypes 
      Height          =   3750
      Left            =   105
      TabIndex        =   5
      Top             =   3090
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   6615
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "LeaveTypesCode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Leave Type"
      Columns(1).DataField=   "LeaveTypesname"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Description"
      Columns(2).DataField=   "Description"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No. of Days"
      Columns(3).DataField=   "LeaveCredits"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Convertible Days"
      Columns(4).DataField=   "ConvertibleDays"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3096"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3016"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4075"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3995"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1746"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1667"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
      HeadLines       =   1
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Named:id=33:Normal"
      _StyleDefs(55)  =   ":id=33,.parent=0"
      _StyleDefs(56)  =   "Named:id=34:Heading"
      _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(58)  =   ":id=34,.wraptext=-1"
      _StyleDefs(59)  =   "Named:id=35:Footing"
      _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   "Named:id=36:Selected"
      _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=37:Caption"
      _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(65)  =   "Named:id=38:HighlightRow"
      _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(67)  =   "Named:id=39:EvenRow"
      _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(69)  =   "Named:id=40:OddRow"
      _StyleDefs(70)  =   ":id=40,.parent=33"
      _StyleDefs(71)  =   "Named:id=41:RecordSelector"
      _StyleDefs(72)  =   ":id=41,.parent=34"
      _StyleDefs(73)  =   "Named:id=42:FilterBar"
      _StyleDefs(74)  =   ":id=42,.parent=33"
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   18
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Edit"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   0
         Left            =   75
         TabIndex        =   19
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&New"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   2
         Left            =   2385
         TabIndex        =   20
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Delete"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   3
         Left            =   3540
         TabIndex        =   21
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Cancel"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   4
         Left            =   4695
         TabIndex        =   22
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Print"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   5
         Left            =   5850
         TabIndex        =   23
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Close"
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
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   7875
      TabIndex        =   24
      Top             =   60
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
End
Attribute VB_Name = "frmMDLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Recordsets
Option Explicit
Dim LeaveTypes As ADODB.Recordset

'Boolean
Dim mAdd As Boolean
Dim mEdit As Boolean
Dim mTransActive As Boolean

'storage
Dim mCode As Integer
Dim mLeaveTypesSortField As String

Private Sub Form_Load()

    Add_MDIButton Me.Name, TitleBar.Caption
        
    Me.Width = 7830
    Me.Height = 7470
    FormCenter Me
    LoadLeaveTypes
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fraButton
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fraSearch
        .Top = fraButton.Top + fraButton.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With frmeLeaveTypes
        .Top = fraSearch.Top + fraSearch.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With gridLeaveTypes
        .Top = frmeLeaveTypes.Top + frmeLeaveTypes.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMDLeave = Nothing
End Sub

Sub FormCenter(Frm As Form)

    Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
    
End Sub

Private Sub LoadLeaveTypes()

    DoEvents
    NetOpen LeaveTypes, "select * from LeaveTypes order by LeaveTypescode"
    DoEvents

    If LeaveTypes.State = adStateOpen Then
        If LeaveTypes.RecordCount > 0 Then
            LeaveTypes.MoveFirst
            Lock_Button "TTTFTT", cmdMenu, 5
        Else
            Lock_Button "TFFFTT", cmdMenu, 5
        End If
        Set gridLeaveTypes.DataSource = LeaveTypes
        mLeaveTypesSortField = "LeaveTypes"
    End If

End Sub

Private Sub cmdmenu_Click(Index As Integer)
'button index procedure
Select Case Index
    Case 0: Add_Record          'execute add record procedure
    Case 1: Edit_Record         'execute edit record procedure
    Case 2: Delete_Record       'execute delete record procedure
    Case 3: Cancel_Transaction  'execute cancel transaction procedure
    Case 4: Print_Record        'execute print record procedure
    Case 5: Close_Form          'execute close form procedure
End Select
End Sub

Private Sub ClearFields()
    txtLeaveTypesCode.Text = "AUTO GENERATED..."
    txtleavetypes.Text = ""
    txtDescription.Text = ""
    chkWpay.Value = 0
    txtLeaveCredits.Value = 0
    txtConvertibleDays.Value = 0
End Sub

Private Sub Add_Record()
If cmdMenu(0).Caption = "&New" Then
    mAdd = True
    mTransActive = True
    cmdMenu(0).Caption = "&Save"
    Lock_Button "TFFTFF", cmdMenu, 5
    frmeLeaveTypes.Enabled = True
    gridLeaveTypes.Enabled = False
    txtSearch.Enabled = False
    Call ClearFields
    SafeSetFocus txtleavetypes
Else
    If txtleavetypes.Text = "" Then
        MsgBox "You need to specify the Leave Type name.", vbInformation
        SafeSetFocus txtleavetypes
        Exit Sub
    End If
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        txtLeaveTypesCode.Text = LastCode("LeaveTypes")
        ConMain.Execute "insert into LeaveTypes values ('" & txtLeaveTypesCode.Text & "', '" & CSQ(txtleavetypes.Text) & "', '" & CSQ(txtDescription.Text) & "', " & IIf(txtLeaveCredits.Value > 0, txtLeaveCredits.Value, 0) & ", " & IIf(txtConvertibleDays.Value > 0, txtConvertibleDays.Value, 0) & ",'" & chkWpay.Value & "')"
    ConMain.CommitTrans
    gridLeaveTypes.Enabled = True
    frmeLeaveTypes.Enabled = False
    txtSearch.Enabled = True
    mCode = txtLeaveTypesCode.Text
    LeaveTypes.Requery
    pointmetdg gridLeaveTypes, LeaveTypes, "LeaveTypesCode", mCode
    mAdd = False
    mTransActive = False
    cmdMenu(0).Caption = "&New"
    Lock_Button "TTTFTT", cmdMenu, 5
End If
End Sub


Private Sub Edit_Record()
If cmdMenu(1).Caption = "&Edit" Then
    mEdit = True
    mTransActive = True
    cmdMenu(1).Caption = "&Save"
    Lock_Button "FTFTFF", cmdMenu, 5
    frmeLeaveTypes.Enabled = True
    gridLeaveTypes.Enabled = False
    txtSearch.Enabled = False
    mCode = txtLeaveTypesCode.Text
    SafeSetFocus txtleavetypes
Else
    If txtleavetypes.Text = "" Then
        MsgBox "You need to specify the Leave Type name Name.", vbInformation
        SafeSetFocus txtleavetypes
        Exit Sub
    End If
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        ConMain.Execute "update LeaveTypes set LeaveTypesname = '" & CSQ(txtleavetypes.Text) & "',  description = '" & CSQ(txtDescription.Text) & "',withpay = '" & chkWpay.Value & "', leavecredits = " & IIf(txtLeaveCredits.Value > 0, txtLeaveCredits.Value, 0) & ", convertibledays = " & IIf(txtConvertibleDays.Value > 0, txtConvertibleDays.Value, 0) & " where LeaveTypescode = '" & txtLeaveTypesCode.Text & "'"
    ConMain.CommitTrans
    gridLeaveTypes.Enabled = True
    frmeLeaveTypes.Enabled = False
    txtSearch.Enabled = True
    LeaveTypes.Requery
    pointmetdg gridLeaveTypes, LeaveTypes, "LeaveTypesCode", mCode
    mEdit = False
    mTransActive = False
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
End If
End Sub

Private Sub Delete_Record()
If LeaveTypes.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        gridLeaveTypes.Delete
    End If
End If
End Sub

Private Sub Cancel_Transaction()
If mAdd = True Then
    cmdMenu(0).Caption = "&New"
    If LeaveTypes.RecordCount > 0 Then
        Lock_Button "TTTFTT", cmdMenu, 5
    Else
        Lock_Button "TFFFTT", cmdMenu, 5
    End If
End If
If mEdit = True Then
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
End If
frmeLeaveTypes.Enabled = False
txtSearch.Enabled = True
gridLeaveTypes.Enabled = True
gridLeaveTypes_RowColChange gridLeaveTypes.Row, gridLeaveTypes.Col
End Sub

Private Sub Print_Record()

End Sub

Private Sub Close_Form()
Unload Me
End Sub

Private Sub gridLeaveTypes_HeadClick(ByVal ColIndex As Integer)
If LeaveTypes.RecordCount > 0 Then
    mLeaveTypesSortField = gridLeaveTypes.Columns(ColIndex).DataField
    LeaveTypes.Sort = mLeaveTypesSortField
End If
End Sub

Private Sub gridLeaveTypes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With LeaveTypes
    If .RecordCount > 0 Then
        txtLeaveTypesCode.Text = !leavetypescode
        txtleavetypes.Text = !LeaveTypesname
        txtDescription.Text = !Description
        chkWpay.Value = !withpay
        txtLeaveCredits.Value = !leavecredits
        txtConvertibleDays.Value = !convertibledays
    Else
        Call ClearFields
    End If
End With
End Sub

Private Sub txtConvertibleDays_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtDescription_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub



Private Sub txtLeaveCredits_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtLeaveTypes_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtLeaveTypesCode_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtSearch_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
Else
    SearchRecord Keyascii, txtSearch, LeaveTypes, txtSearch.Text, mLeaveTypesSortField
End If
End Sub


