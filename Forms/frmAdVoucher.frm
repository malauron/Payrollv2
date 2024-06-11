VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAdVoucher 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   14730
   Tag             =   "Create Voucher(s)"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14730
      TabIndex        =   8
      Top             =   0
      Width           =   14730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Create Voucher(s)"
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
         Top             =   225
         Width           =   5445
      End
   End
   Begin VB.Frame fraBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   120
      TabIndex        =   6
      Top             =   8835
      Width           =   10815
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   480
         Left            =   135
         TabIndex        =   7
         Top             =   0
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   847
         Caption         =   "Clo&se"
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
         Image           =   "frmAdVoucher.frx":0000
         cBack           =   14737632
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   8115
      Left            =   15
      TabIndex        =   10
      Top             =   645
      Width           =   14250
      Begin VB.Frame fraMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   6060
         TabIndex        =   18
         Top             =   5910
         Width           =   8145
         Begin VB.OptionButton optButcheryPastry 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Butchery Pastry Voucher"
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
            Left            =   2745
            TabIndex        =   23
            Top             =   60
            Width           =   2565
         End
         Begin VB.OptionButton optGroceryVoucher 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Grocery Voucher"
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
            Left            =   255
            TabIndex        =   22
            Top             =   60
            Value           =   -1  'True
            Width           =   1875
         End
         Begin lvButton.lvButtons_H cmdCreateVoucher 
            Height          =   480
            Left            =   6480
            TabIndex        =   5
            Top             =   270
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   847
            Caption         =   "C&reate"
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
            Image           =   "frmAdVoucher.frx":0CDA
            cBack           =   14737632
         End
         Begin TDBNumber6Ctl.TDBNumber txtAmount 
            Height          =   315
            Left            =   90
            TabIndex        =   1
            Top             =   615
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            Calculator      =   "frmAdVoucher.frx":19B4
            Caption         =   "frmAdVoucher.frx":19D4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdVoucher.frx":1A40
            Keys            =   "frmAdVoucher.frx":1A5E
            Spin            =   "frmAdVoucher.frx":1AA8
            AlignHorizontal =   1
            AlignVertical   =   2
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
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBDate6Ctl.TDBDate txtDateIssued 
            Height          =   315
            Left            =   1380
            TabIndex        =   2
            Top             =   615
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            Calendar        =   "frmAdVoucher.frx":1AD0
            Caption         =   "frmAdVoucher.frx":1BD6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdVoucher.frx":1C3C
            Keys            =   "frmAdVoucher.frx":1C5A
            Spin            =   "frmAdVoucher.frx":1CB8
            AlignHorizontal =   2
            AlignVertical   =   2
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
         Begin TDBDate6Ctl.TDBDate txtStartDate 
            Height          =   315
            Left            =   2670
            TabIndex        =   3
            Top             =   615
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            Calendar        =   "frmAdVoucher.frx":1CE0
            Caption         =   "frmAdVoucher.frx":1DE6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdVoucher.frx":1E4C
            Keys            =   "frmAdVoucher.frx":1E6A
            Spin            =   "frmAdVoucher.frx":1EC8
            AlignHorizontal =   2
            AlignVertical   =   2
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
         Begin TDBDate6Ctl.TDBDate txtEndDate 
            Height          =   315
            Left            =   4005
            TabIndex        =   4
            Top             =   615
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            Calendar        =   "frmAdVoucher.frx":1EF0
            Caption         =   "frmAdVoucher.frx":1FF6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdVoucher.frx":205C
            Keys            =   "frmAdVoucher.frx":207A
            Spin            =   "frmAdVoucher.frx":20D8
            AlignHorizontal =   2
            AlignVertical   =   2
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Validity Date"
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
            Left            =   2670
            TabIndex        =   21
            Top             =   420
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Issued"
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
            Left            =   1380
            TabIndex        =   20
            Top             =   420
            Width           =   1200
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
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
            Left            =   90
            TabIndex        =   19
            Top             =   420
            Width           =   1200
         End
      End
      Begin VB.Frame fraClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   90
         TabIndex        =   16
         Top             =   6330
         Width           =   6540
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   480
            Left            =   45
            TabIndex        =   17
            Top             =   60
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   847
            Caption         =   "&Cancel"
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
            Image           =   "frmAdVoucher.frx":2100
            cBack           =   14737632
         End
      End
      Begin TrueOleDBGrid80.TDBGrid tdgLoan 
         Height          =   5475
         Left            =   105
         TabIndex        =   11
         Top             =   795
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   9657
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Voucher Number"
         Columns(0).DataField=   "dummycode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Amount"
         Columns(1).DataField=   "amount"
         Columns(1).NumberFormat=   "#,##0.00"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Date Issued"
         Columns(2).DataField=   "dateissued"
         Columns(2).NumberFormat=   "YYYY-MM-DD"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Status"
         Columns(3).DataField=   "status"
         Columns(3).NumberFormat=   "#,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3043"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2963"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2355"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerStyle=0"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2302"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(3)._HeadDivider=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=79"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=102,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=99,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=100,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin TDBText6Ctl.TDBText txtFullname 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   210
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   529
         Caption         =   "frmAdVoucher.frx":2DDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdVoucher.frx":2E46
         Key             =   "frmAdVoucher.frx":2E64
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
      Begin lvButton.lvButtons_H cmdSearchEmployee 
         Height          =   315
         Left            =   6270
         TabIndex        =   12
         ToolTipText     =   "Browse for checked in guests."
         Top             =   210
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Caption         =   "..."
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
         cBack           =   14737632
      End
      Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer xrpt 
         Height          =   5475
         Left            =   6690
         TabIndex        =   15
         Top             =   420
         Width           =   7515
         lastProp        =   600
         _cx             =   13256
         _cy             =   9657
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
         Left            =   -495
         TabIndex        =   14
         Top             =   255
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "History"
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
         Left            =   105
         TabIndex        =   13
         Top             =   585
         Width           =   6525
      End
   End
End
Attribute VB_Name = "frmAdVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mBranchCode              As String
Public mDivisionCode            As String
Public mCostCenterCode          As String

Public mEmployeeCode            As Integer

Public mNew                     As Boolean
Public mContinue                As Boolean

Public rsVoucher                As ADODB.Recordset

Dim CrxRep                      As CRAXDRT.Report

Private Sub cmdCancel_Click()
    
    Dim mVoucherCode            As String
    
    If MsgBox("Do you want to cancel this voucher?", vbQuestion + vbYesNo) = vbYes Then
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        ConMain.Execute "update voucher set status = 'Cancelled' where vouchercode = " & rsVoucher!vouchercode & ""
        mVoucherCode = rsVoucher!vouchercode
        ConMain.CommitTrans
        
        rsVoucher.Requery
    
        rsVoucher.MoveFirst
        rsVoucher.Find "vouchercode = " & mVoucherCode & ""
    
    End If
    
End Sub

Private Sub cmdClose_Click()
  
  Unload Me
  
End Sub

Private Sub cmdCreateVoucher_Click()
    
    Dim mVoucherCode        As Integer
    
    Dim mVoucherType        As String
    
    If mEmployeeCode < 1 Then
        MsgBox "Please select an employee.", vbExclamation + vbOKOnly
        cmdSearchEmployee.SetFocus
        Exit Sub
    End If
    
    If CDbl(txtAmount.Text) <= 0 Then
        MsgBox "Please enter an amount greater than zero.", vbExclamation + vbOKOnly
        txtAmount.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDateIssued.Text) Then
        MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
        txtDateIssued.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtStartDate.Text) Then
        MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
        txtStartDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtEndDate.Text) Then
        MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
        txtEndDate.SetFocus
        Exit Sub
    End If
    
    If CDate(txtEndDate.Text) < CDate(txtStartDate.Text) Then
        MsgBox "Ending date of validity is earlier than its start date.", vbExclamation + vbOKOnly
        txtStartDate.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Creating new voucher. Do you want to proceed?", vbQuestion + vbYesNo) = vbYes Then
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        
        'mVoucherCode = LastCode("Voucher")
        
        If optGroceryVoucher.Value = True Then
            mVoucherType = "groc"
            mVoucherCode = LastCode("Grocery Voucher")
        Else
            mVoucherType = "butc"
            mVoucherCode = LastCode("Butchery Voucher")
        End If
        
        ConMain.Execute "insert into voucher (vouchercode,dummycode,employeecode,amount,dateissued,startdate,enddate,vouchertype) values " & _
                        "(" & mVoucherCode & ",'" & Format(mVoucherCode, "0000000000-" & mVoucherType) & "'," & mEmployeeCode & "," & Format(txtAmount.Text, "##0.00") & ",'" & Format(txtDateIssued.Text, "YYYY-MM-DD") & "','" & Format(txtStartDate.Text, "YYYY-MM-DD") & "','" & Format(txtEndDate.Text, "YYYY-MM-DD") & "','" & mVoucherType & "')"
        
        ConMain.CommitTrans
        rsVoucher.Requery
        rsVoucher.MoveFirst
        rsVoucher.Find "dummycode = '" & Format(mVoucherCode, "0000000000-" & mVoucherType) & "'"
        
        Print_Voucher mVoucherCode, mVoucherType
        
        txtAmount.Text = "0.00"
        txtDateIssued.Text = Format(Now, "MM/DD/YYYY")
        txtStartDate.Text = ""
        txtEndDate.Text = ""
    
    End If
        
End Sub
    
Private Sub cmdSearchEmployee_Click()
    With frmBrowseEmployee
        txtFullname.SetFocus
        .mBrowseType = "Voucher"
        .Show vbModal
    End With
End Sub
    
Private Sub Form_Activate()
    
    Focus_MDIButton Me
    
End Sub
    
Private Sub Form_Load()

    Add_MDIButton Me.Name, Me.Tag
    
    NetOpen rsVoucher, "select * from voucher where vouchercode = 0"
    
    Set tdgLoan.DataSource = rsVoucher
    
    txtAmount.Text = "0.00"
    txtDateIssued.Text = Format(Now, "MM/DD/YYYY")
    txtStartDate.Text = ""
    txtEndDate.Text = ""
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Remove_MDIButton Me.Name
    
    mBranchCode = 0
    mDivisionCode = 0
    mCostCenterCode = 0
    
    mEmployeeCode = 0
    
    mNew = False
    mContinue = False
    
    Set rsVoucher = Nothing

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With fraMain
        .Top = (pic1.Top + pic1.Height) - 70
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - (pic1.Height + fraBottom.Height)
    End With

    With fraBottom
        .Top = fraMain.Top + fraMain.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With tdgLoan
        .Height = fraMain.Height - (.Top + fraClose.Height + 100)
    End With
    
    With fraClose
        .Top = tdgLoan.Top + tdgLoan.Height + 25
    End With
    
    With xrpt
        .Top = tdgLoan.Top
        .Width = fraMain.Width - (.Left + 100)
        .Height = (tdgLoan.Height + fraClose.Height) - (fraMenu.Height)
    End With
    
    With fraMenu
        .Top = xrpt.Top + xrpt.Height + 25
        .Left = xrpt.Left
        .Width = xrpt.Width
    End With
    
End Sub

Private Sub tdgLoan_DblClick()
    
    With rsVoucher
        If .RecordCount > 0 Then
            If !Status <> "Cancelled" Then
                Print_Voucher !vouchercode, !vouchertype
            End If
        End If
    End With
    
End Sub

Private Sub tdgLoan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With rsVoucher
        If .RecordCount > 0 Then
            If Not .EOF Then
                If !Status <> "Cancelled" Then
                    cmdCancel.Enabled = True
                Else
                    cmdCancel.Enabled = False
                End If
            Else
                cmdCancel.Enabled = False
            End If
        Else
            cmdCancel.Enabled = False
        End If
    End With
End Sub
    
Private Sub txtamount_GotFocus()
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)
End Sub
    
Private Sub txtamount_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub txtDateIssued_GotFocus()
    txtDateIssued.SelStart = 0
    txtDateIssued.SelLength = Len(txtDateIssued.Text)
End Sub
    
Private Sub txtDateIssued_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub txtEndDate_GotFocus()
    txtEndDate.SelStart = 0
    txtEndDate.SelLength = Len(txtEndDate.Text)
End Sub
    
Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub txtFullname_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub txtStartdate_GotFocus()
    txtStartDate.SelStart = 0
    txtStartDate.SelLength = Len(txtStartDate.Text)
End Sub
    
Private Sub txtStartdate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub Print_Voucher(mVoucherCode As Integer, mVoucherType As String)
    
    Dim CrxApp              As CRAXDRT.Application
    Dim crxDatabase         As CRAXDRT.Database
    Dim crxDatabaseTable    As CRAXDRT.DatabaseTable
    Dim crxDatabaseTables   As CRAXDRT.DatabaseTables
    
    Set CrxApp = Nothing
    Set CrxRep = Nothing
    
    Set CrxApp = New CRAXDRT.Application
    Set CrxRep = New CRAXDRT.Report
        
    Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\rptVoucher.rpt")
    
    Set crxDatabase = CrxRep.Database
    Set crxDatabaseTables = crxDatabase.Tables
    
    For Each crxDatabaseTable In crxDatabaseTables
        crxDatabaseTable.ConnectionProperties("connection string") = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & SQLServerName & "; DATABASE=" & SQLDatabase & "; UID=" & SQLUsername & "; PWD=" & SQLPassword & "; PORT='" & SQLPort & "'"
    Next crxDatabaseTable
    
    CrxRep.ParameterFields.GetItemByName("mVoucherCode").AddCurrentValue mVoucherCode
    CrxRep.ParameterFields.GetItemByName("mVoucherType").AddCurrentValue mVoucherType
    
    xrpt.ReportSource = CrxRep
    xrpt.ViewReport
    xrpt.Zoom 100
    
    Set crxDatabase = Nothing
    Set crxDatabaseTable = Nothing
    Set crxDatabaseTables = Nothing
    Set CrxApp = Nothing
    
End Sub

Private Sub xrpt_PrintButtonClicked(UseDefault As Boolean)
    CrxRep.PrinterSetup Me.hwnd
End Sub

