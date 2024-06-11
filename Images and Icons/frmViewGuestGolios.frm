VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmViewGuestGolios 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   15240
   Tag             =   "View Guest Folio"
   WindowState     =   2  'Maximized
   Begin VB.Frame fraBedsReserved 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reserved Beds"
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
      Left            =   8880
      TabIndex        =   0
      Top             =   1710
      Width           =   6810
      Begin VB.Frame fraButton2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   45
         TabIndex        =   1
         Top             =   4125
         Width           =   7125
         Begin lvButton.lvButtons_H cmdCancelBedsReserved 
            Height          =   480
            Left            =   45
            TabIndex        =   2
            Top             =   45
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   847
            Caption         =   "&CANCEL"
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
            Image           =   "frmViewGuestGolios.frx":0000
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   480
            Left            =   60
            TabIndex        =   3
            Top             =   555
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   847
            Caption         =   "&Register"
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
            Image           =   "frmViewGuestGolios.frx":0CDA
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin VB.Label lblIndividual 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Numbers Beds Reserved For "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2295
            TabIndex        =   4
            Top             =   540
            Width           =   6300
         End
      End
      Begin TrueOleDBGrid80.TDBGrid tdgBedsReserved 
         Height          =   3885
         Left            =   45
         TabIndex        =   5
         Top             =   180
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "include"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Charges"
         Columns(1).DataField=   "chargename"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Qty."
         Columns(2).DataField=   "quantity"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Price"
         Columns(3).DataField=   "chargeprice"
         Columns(3).NumberFormat=   "#,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Gross Amt."
         Columns(4).DataField=   "gross_amount"
         Columns(4).NumberFormat=   "#,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Discount"
         Columns(5).DataField=   "discount"
         Columns(5).NumberFormat=   "#,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Net Amount"
         Columns(6).DataField=   "net_amount"
         Columns(6).NumberFormat=   "#,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   "chargecode"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8705"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2778"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2699"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8705"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1349"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1270"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8706"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8706"
         Splits(0)._ColumnProps(25)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=2223"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2143"
         Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(33)=   "Column(5).Width=2355"
         Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2275"
         Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(39)=   "Column(6).Width=1191"
         Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=1111"
         Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(45)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=8708"
         Splits(0)._ColumnProps(50)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00808080&
      ForeColor       =   &H00404040&
      Height          =   675
      Left            =   15
      TabIndex        =   16
      Top             =   750
      Width           =   13620
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   9015
         TabIndex        =   17
         Top             =   225
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   529
         Caption         =   "frmViewGuestGolios.frx":1454
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmViewGuestGolios.frx":14C0
         Key             =   "frmViewGuestGolios.frx":14DE
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
         Left            =   5310
         TabIndex        =   18
         Tag             =   "Municipal"
         Top             =   195
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
         _PropDict       =   $"frmViewGuestGolios.frx":1522
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
      Begin TrueOleDBList80.TDBCombo tdbShowCriteria 
         Height          =   345
         Left            =   780
         TabIndex        =   19
         Tag             =   "Municipal"
         Top             =   180
         Width           =   3705
         _ExtentX        =   6535
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
         _PropDict       =   $"frmViewGuestGolios.frx":15CC
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW"
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
         Left            =   -300
         TabIndex        =   22
         Top             =   255
         Width           =   1005
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
         Left            =   7950
         TabIndex        =   21
         Top             =   255
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
         Left            =   4230
         TabIndex        =   20
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame fraReservation 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reservation Entries"
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
      Left            =   0
      TabIndex        =   10
      Top             =   1710
      Width           =   8895
      Begin VB.Frame fraButton1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   30
         TabIndex        =   11
         Top             =   5115
         Width           =   8805
         Begin lvButton.lvButtons_H cmdCancelReservation 
            Height          =   480
            Left            =   2070
            TabIndex        =   12
            Top             =   45
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   847
            Caption         =   "&CANCEL"
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
            Image           =   "frmViewGuestGolios.frx":1676
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdRegister 
            Height          =   480
            Left            =   60
            TabIndex        =   13
            Top             =   45
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   847
            Caption         =   "&Check Out"
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
            Image           =   "frmViewGuestGolios.frx":2350
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin VB.Label lblAll 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Reservations"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4485
            TabIndex        =   14
            Top             =   600
            Width           =   6285
         End
      End
      Begin TrueOleDBGrid80.TDBGrid tdgRegistration 
         Height          =   3885
         Left            =   30
         TabIndex        =   15
         Top             =   180
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Registration Code"
         Columns(0).DataField=   "dummycode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "OR Number"
         Columns(1).DataField=   "ornumber"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Guest"
         Columns(2).DataField=   "fullname"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Arrival Date"
         Columns(3).DataField=   "arrivaldate"
         Columns(3).NumberFormat=   "MM-DD-YYYY"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Departure Date"
         Columns(4).DataField=   "departuredate"
         Columns(4).NumberFormat=   "MM-DD-YYYY"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Log Status"
         Columns(5).DataField=   "logstatus"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2275"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8705"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5265"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5186"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8708"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2249"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2170"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2302"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2223"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=873"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=794"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
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
      ScaleWidth      =   15240
      TabIndex        =   8
      Top             =   0
      Width           =   15240
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "View/Manage Guest Folios"
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
         Width           =   4410
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   360
      TabIndex        =   6
      Top             =   8115
      Width           =   7410
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   465
         Left            =   60
         TabIndex        =   7
         Top             =   45
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   820
         Caption         =   "&CLOSE"
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
         Image           =   "frmViewGuestGolios.frx":302A
         cBack           =   14737632
      End
   End
End
Attribute VB_Name = "frmViewGuestGolios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsRegistration           As ADODB.Recordset
Dim rsChargeFileTmp       As ADODB.Recordset
'Dim rsReservedBedCount      As ADODB.Recordset

Private Sub cmdCancelBedsReserved_Click()
        
    Dim mCount      As Integer
    
    mCount = 0
    With rsChargeFileTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !include <> 0 Then
                    mCount = mCount + 1
                End If
                .MoveNext
            Loop
            If mCount = 0 Then
                MsgBox "Please select a bed.", vbExclamation + vbOKOnly
                Exit Sub
            End If
            If mCount = .RecordCount Then
                cmdCancelReservation_Click
            Else
                ConMain.Execute "set autocommit = 0"
                ConMain.BeginTrans
                .MoveFirst
                Do While Not .EOF
                    If !include <> 0 Then
                        
                        If Format(rsRegistration!arrivaldate, "MM/DD/YYYY") <= Format(Now, "MM/DD/YYYY") And Format(rsRegistration!departuredate, "MM/DD/YYYY") >= Format(Now, "MM/DD/YYYY") Then
                            ConMain.Execute "Update bed set status = 2, currentguest = " & !guestcode & " where bedcode = " & !bedcode & ""
                        End If
                        
                        ConMain.Execute "insert into cancelledbedregistration(registrationcode,floorcode,sectioncode,bedlevelcode," & _
                                                        "bedcode,usercode,trnxdate,trnxtime) values ( " & _
                                                        rsRegistration!registrationcode & "," & !floorcode & "," & !sectioncode & "," & !bedlevelcode & "," & _
                                                        !bedcode & "," & UserCode & ",curdate(),curtime())"
                        ConMain.Execute "update bedsreserved set status = 'Cancelled' where registrationcode = " & rsRegistration!registrationcode & " and bedcode = " & !bedcode & ""
                        ConMain.Execute "delete from dailybedstatus where bedcode = " & !bedcode & " and " & _
                                            "datetaken between '" & Format(rsRegistration!arrivaldate, "YYYY-MM-DD") & "' and '" & Format(rsRegistration!departuredate, "YYYY-MM-DD") & "' "
                    End If
                    .MoveNext
                Loop
                ConMain.CommitTrans
                'rsReservedBedCount.Requery
                'lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
                tdgregistration_DblClick
            End If
        Else
            MsgBox "Please select a bed.", vbExclamation + vbOKOnly
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCancelReservation_Click()
    
    Dim rsTmp       As ADODB.Recordset
    
    
    With rsRegistration
        
        If !logstatus <> "In" Then
            MsgBox "You can not cancell this registration.", vbExclamation + vbOKOnly
            Exit Sub
        End If
        
        If .RecordCount > 0 Then
            If Not .EOF Then
                If MsgBox("Do you want to cancel this registration?", vbQuestion + vbYesNo) = vbYes Then
                    ConMain.Execute "set autocommit = 0 "
                    ConMain.BeginTrans
                    
                    ConMain.Execute "update registration set logstatus = 'Out' where registrationcode = " & !registrationcode & ""
                    
                    If Format(rsRegistration!arrivaldate, "MM/DD/YYYY") <= Format(Now, "MM/DD/YYYY") And Format(rsRegistration!departuredate, "MM/DD/YYYY") >= Format(Now, "MM/DD/YYYY") Then
                        ConMain.Execute "Update bed set status = 2, currentguest = " & !guestcode & " where bedcode = " & !bedcode & ""
                    End If
                    ConMain.Execute "insert into cancelledbedregistration(registrationcode,floorcode,sectioncode,bedlevelcode," & _
                                                "bedcode,usercode,trnxdate,trnxtime) values ( " & _
                                                !registrationcode & "," & !floorcode & "," & !sectioncode & "," & !bedlevelcode & "," & _
                                                !bedcode & "," & UserCode & ",curdate(),curtime())"
                    ConMain.Execute "delete from dailybedstatus where bedcode = " & !bedcode & " and " & _
                                    "datetaken between '" & Format(!arrivaldate, "YYYY-MM-DD") & "' and '" & Format(!departuredate, "YYYY-MM-DD") & "' "
                    
                    'ConMain.Execute "update bedsreserved set status = 'Cancelled' where registrationcode = " & !registrationcode & ""
                    ConMain.CommitTrans
                    rsRegistration.Requery
                    'rsReservedBedCount.Requery
                    'lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
                    'lblAll.Caption = "Total Registrations : " & rsRegistration.RecordCount
                    Create_ItemsAvailedTmp
                End If
            End If
        Else
            MsgBox "No registration to cancel.", vbExclamation + vbOKOnly
        End If
        
    End With
    
End Sub

Private Sub cmdRegister_Click()

    Dim mCode                   As Integer
    
    Dim mCompanyCode            As String
    
    Dim mtotal                  As Double
    
    Dim rsGuestTmp              As ADODB.Recordset
    Dim rsChargeFile          As ADODB.Recordset
    
    If rsRegistration.RecordCount > 0 Then
    
        If CDate(Format(rsRegistration!arrivaldate, "MM/DD/YYYY")) > CDate(Format(Now, "MM/DD/YYYY")) Then
            
        End If

        NetOpen rsChargeFile, "select x1.* from bedsreserved x1 where x1.registrationcode = " & rsRegistration!registrationcode & " and x1.status = 'For Registration'"
        
        With rsChargeFile
        
            If .RecordCount > 0 Then
                
                ConMain.Execute "set autocommit = 0"
                ConMain.BeginTrans
                    
                ConMain.Execute "update reservation set status = 'Registered' where registrationcode = " & rsRegistration!registrationcode & ""
                
                .MoveFirst
                
                Do While Not .EOF
                
                    mCode = LastCode("Registration")
                
                    NetOpen rsGuestTmp, "select * from guest where guestcode = " & !guestcode & ""
                    
                    If IsNull(rsGuestTmp!companycode) Then
                        mCompanyCode = "Null"
                    Else
                        mCompanyCode = "'" & rsGuestTmp!companycode & "'"
                    End If
                    
                    mtotal = !bedrate * rsRegistration!noofdays
                    
                    ConMain.Execute "insert into registration (registrationcode,guestcode,modeofpayment,chargeto,companycode, " & _
                                    "sectioncode,bedlevelcode,bedcode,arrivaldate,arrivaltime, " & _
                                    "departuredate,departuretime,bedrate,total, " & _
                                    "baggageqty,baggagettlfee,ornumber,linenqty,linentotalprice, " & _
                                    "othersqty,otherstotalprice,logstatus,floorcode,noofdays," & _
                                    "registrationcode,usercode,trnxdate,trnxtime,dummycode) values  " & _
                                    "(" & mCode & "," & rsGuestTmp!guestcode & ",'" & rsGuestTmp!modeofpayment & "','" & rsGuestTmp!chargeto & "'," & mCompanyCode & ", " & _
                                    !sectioncode & "," & !bedlevelcode & "," & !bedcode & ",'" & Format(rsRegistration!arrivaldate, "YYYY-MM-DD") & "','" & Format(rsRegistration!arrivaltime, "HH:NN:SS") & "','" & Format(rsRegistration!departuredate, "YYYY-MM-DD") & "','" & Format(rsRegistration!departuretime, "HH:NN:SS") & "'," & Format(!bedrate, "##0.00") & "," & Format(mtotal, "##0.00") & "," & _
                                    0 & "," & 0 & ",''," & 0 & "," & 0 & ", " & _
                                    0 & "," & 0 & ",'In'," & !floorcode & "," & rsRegistration!noofdays & "," & _
                                    rsRegistration!registrationcode & "," & UserCode & ",curdate(),curtime(),'" & Format(mCode, "0000000000") & "')"
                    
                    ConMain.Execute "Update bed set status = 1, currentguest = " & rsGuestTmp!guestcode & " where bedcode = " & !bedcode & ""
                    
                    ConMain.Execute "insert into eventslog (modulename,actiontaken,trnxdate,usercode,username) values " & _
                                      "('" & Me.Name & "','Save Registration- " & mCode & "', Now()," & UserCode & ",'" & UserName & "')"
                            
                    ConMain.Execute "update bedsreserved set status = 'Registered' where registrationcode = " & rsRegistration!registrationcode & " and bedcode = " & !bedcode & ""
                    ConMain.Execute "update dailybedstatus set status = 1 where bedcode = " & !bedcode & " and " & _
                                        "datetaken between '" & Format(rsRegistration!arrivaldate, "YYYY-MM-DD") & "' and '" & Format(rsRegistration!departuredate, "YYYY-MM-DD") & "' "
                    .MoveNext
                    
                Loop
                
                ConMain.CommitTrans
                
                rsRegistration.Requery
'                rsReservedBedCount.Requery
'                lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
                lblAll.Caption = "Total Registrations : " & rsRegistration.RecordCount
                tdgregistration_DblClick
                
            End If
            
        End With
            
    End If
    
End Sub

Private Sub cmdSave_Click()


    Dim mCode           As Integer
    Dim mCount          As Integer
    
    Dim mtotal          As Double
    
    Dim mCompanyCode    As String
    
    Dim rsGuestTmp      As ADODB.Recordset
    
    mCount = 0
    With rsChargeFileTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !include <> 0 Then
                    mCount = mCount + 1
                End If
                .MoveNext
            Loop
            If mCount = 0 Then
                MsgBox "Please select a bed.", vbExclamation + vbOKOnly
                Exit Sub
            End If
            ConMain.Execute "set autocommit = 0"
            ConMain.BeginTrans
            
            If mCount = .RecordCount Then
                ConMain.Execute "update reservation set status = 'Registered' where registrationcode = " & rsRegistration!registrationcode & ""
            End If
            
            .MoveFirst
            
            Do While Not .EOF
                If !include <> 0 Then
                
                
                    mCode = LastCode("Registration")
                
                    NetOpen rsGuestTmp, "select * from guest where guestcode = " & !guestcode & ""
                    
                    If IsNull(rsGuestTmp!companycode) Then
                        mCompanyCode = "Null"
                    Else
                        mCompanyCode = "'" & rsGuestTmp!companycode & "'"
                    End If
                    
                    mtotal = !bedrate * !noofdays
                    
                    ConMain.Execute "insert into registration (registrationcode,guestcode,modeofpayment,chargeto,companycode, " & _
                                    "sectioncode,bedlevelcode,bedcode,arrivaldate,arrivaltime, " & _
                                    "departuredate,departuretime,bedrate,total, " & _
                                    "baggageqty,baggagettlfee,ornumber,linenqty,linentotalprice, " & _
                                    "othersqty,otherstotalprice,logstatus,floorcode,noofdays," & _
                                    "registrationcode,usercode,trnxdate,trnxtime,dummycode) values  " & _
                                    "(" & mCode & "," & rsGuestTmp!guestcode & ",'" & rsGuestTmp!modeofpayment & "','" & rsGuestTmp!chargeto & "'," & mCompanyCode & ", " & _
                                    !sectioncode & "," & !bedlevelcode & "," & !bedcode & ",'" & Format(rsRegistration!arrivaldate, "YYYY-MM-DD") & "','" & Format(rsRegistration!arrivaltime, "HH:NN:SS") & "','" & Format(rsRegistration!departuredate, "YYYY-MM-DD") & "','" & Format(rsRegistration!departuretime, "HH:NN:SS") & "'," & Format(!bedrate, "##0.00") & "," & Format(!bedrate, "##0.00") & "," & _
                                    0 & "," & 0 & ",''," & 0 & "," & 0 & ", " & _
                                    0 & "," & 0 & ",'In'," & !floorcode & "," & rsRegistration!noofdays & "," & _
                                    rsRegistration!registrationcode & "," & UserCode & ",curdate(),curtime(),'" & Format(mCode, "0000000000") & "')"
                    
                    ConMain.Execute "Update bed set status = 1, currentguest = " & rsGuestTmp!guestcode & " where bedcode = " & !bedcode & ""
                    
                    ConMain.Execute "insert into eventslog (modulename,actiontaken,trnxdate,usercode,username) values " & _
                                      "('" & Me.Name & "','Save Registration- " & mCode & "', Now()," & UserCode & ",'" & UserName & "')"
                            
                            
                    
                    
                        ConMain.Execute "update bedsreserved set status = 'Registered' where registrationcode = " & rsRegistration!registrationcode & " and bedcode = " & !bedcode & ""
                        ConMain.Execute "update dailybedstatus set status = 1 where bedcode = " & !bedcode & " and " & _
                                            "datetaken between '" & Format(rsRegistration!arrivaldate, "YYYY-MM-DD") & "' and '" & Format(rsRegistration!departuredate, "YYYY-MM-DD") & "' "
                End If
                .MoveNext
            Loop
            
            ConMain.CommitTrans
            
            If mCount = .RecordCount Then
                rsRegistration.Requery
            End If
            
'            rsReservedBedCount.Requery
'            lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
            lblAll.Caption = "Total Registrations : " & rsRegistration.RecordCount
            tdgregistration_DblClick
            
        Else
            
            MsgBox "Please select a bed.", vbExclamation + vbOKOnly
            
        End If
    End With
        
End Sub

Private Sub Form_Activate()

    Focus_MDIButton Me
    
End Sub

Private Sub Form_Load()

    Dim rsTmp           As ADODB.Recordset
    Dim I               As Integer
    
    Add_MDIButton Me.Name, Me.Tag
    
    CreateTmpDB rsTmp
    
    With rsTmp
        For I = 1 To 2
            .AddNew
            If I = 1 Then
                .Fields("code") = "Today's Departure"
                .Fields("description") = "Today's Departure"
            End If
            If I = 2 Then
                .Fields("code") = "All Registrations"
                .Fields("description") = "All Registrations"
            End If
            .Update
        Next
    End With
    
    With tdbShowCriteria
        .RowSource = rsTmp
        .ListField = "description"
        .BoundColumn = "code"
        .Columns(0).DataField = "code"
        .Columns(1).DataField = "description"
        .BoundText = "Today's Departure"
    End With
    
    CreateTmpDB rsTmp
    
    With tdgRegistration
        For I = .Columns("dummycode").ColIndex To .Columns("departuredate").ColIndex
            If .Columns(I).Visible = True Then
                rsTmp.AddNew
                rsTmp.Fields("code") = .Columns(I).DataField
                rsTmp.Fields("description") = .Columns(I).Caption
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
        .BoundText = "dummycode"
    End With
    
    Set rsTmp = Nothing
    
    NetOpen rsRegistration, "select  x1.*," & _
                          "concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname from registration x1 " & _
                          "left outer join guest x2 on x1.guestcode = x2.guestcode " & _
                          "where x1.logstatus = 'In' and x1.departuredate = '" & Format(Now, "MM/DD/YYYY") & "' order by x1.registrationcode"
                          
    Set tdgRegistration.DataSource = rsRegistration
    
    'NetOpen rsReservedBedCount, "select * from bedsreserved where status = 'For Registration'"
    
    'lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
    lblAll.Caption = "Total Registrations : " & rsRegistration.RecordCount
    
    tdgregistration_DblClick
    
    tdgregistration_RowColChange 0, 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Remove_MDIButton Me.Name

End Sub


Private Sub Form_Resize()

    On Error Resume Next
    
    With fraSearch
        .Top = pic1.Top + pic1.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fraReservation
        .Top = fraSearch.Top + fraSearch.Height
        .Left = 0
        .Height = Me.ScaleHeight - (.Top + fraButtons.Height)
    End With
    
    With tdgRegistration
        .Height = fraReservation.Height - 900
    End With
    
    With fraButton1
        .Top = tdgRegistration.Top + tdgRegistration.Height + 50
    End With
    
    With fraButtons
        .Top = fraReservation.Top + fraReservation.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fraBedsReserved
        .Top = fraSearch.Top + fraSearch.Height
        .Left = fraReservation.Left + fraReservation.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - (.Top + fraButtons.Height)
    End With
    
    With tdgBedsReserved
        .Left = 30
        .Width = fraBedsReserved.Width - 90
        .Height = fraReservation.Height - 900
    End With
    
     With fraButton2
        .Left = 30
        .Width = fraBedsReserved.Width - 90
        .Top = tdgBedsReserved.Top + tdgBedsReserved.Height + 50
    End With
    
End Sub

Private Sub tdbSearch_ItemChange()
    rsRegistration.Sort = tdbSearch.BoundText
End Sub

Private Sub tdbSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbSearch, tdbSearch.RowSource, tdbSearch.Text
        tdbSearch_ItemChange
    End If
End Sub

Private Sub tdbShowCriteria_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If tdbShowCriteria.BoundText = "All Registrations" Then
            
            NetOpen rsRegistration, "select  x1.*," & _
                          "concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname from registration x1 " & _
                          "left outer join guest x2 on x1.guestcode = x2.guestcode " & _
                          "order by x1.registrationcode"
        Else
            NetOpen rsRegistration, "select  x1.*," & _
                          "concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname from registration x1 " & _
                          "left outer join guest x2 on x1.guestcode = x2.guestcode " & _
                          "where x1.logstatus = 'In' and x1.departuredate = '" & Format(Now, "MM/DD/YYYY") & "' order by x1.registrationcode"
        End If
        
        Set tdgRegistration.DataSource = rsRegistration
        
        'NetOpen rsReservedBedCount, "select * from bedsreserved where status = 'For Registration'"
        'lblAll.Caption = "Total Number of Beds Reserved : " & rsReservedBedCount.RecordCount
        lblAll.Caption = "Total Registrations : " & rsRegistration.RecordCount
        
        'tdgregistration_DblClick
        
        SendKeys "{TAB}"
        
    Else
        SearchList KeyAscii, tdbShowCriteria, tdbShowCriteria.RowSource, tdbShowCriteria.Text
    End If
End Sub

Private Sub tdgregistration_DblClick()
    
    Dim rsChargeFile          As ADODB.Recordset
    
    If rsRegistration.RecordCount > 0 Then

        NetOpen rsChargeFile, "select x1.*,x2.chargename from chargefile x1 " & _
                              "left outer join charge x2 on x1.chargecode = x2.chargecode " & _
                              "where registrationcode = " & rsRegistration!registrationcode & ""
       With rsChargeFile
            
            If .RecordCount > 0 Then
                'lblIndividual.Caption = "No. of beds reserved for " & !fullname & ": " & .RecordCount
                Create_ItemsAvailedTmp
                .MoveFirst
                Do While Not .EOF
                    rsChargeFileTmp.AddNew
                    rsChargeFileTmp.Fields("chargecode") = !chargecode
                    rsChargeFileTmp.Fields("chargename") = !chargename
                    rsChargeFileTmp.Fields("quantity") = !quantity
                    rsChargeFileTmp.Fields("chargeprice") = !chargeprice
                    rsChargeFileTmp.Fields("gross_amount") = !gross_amount
                    rsChargeFileTmp.Fields("discount") = !discount
                    rsChargeFileTmp.Fields("net_amount") = !net_amount
                    rsChargeFileTmp.Update
                    .MoveNext
                Loop
            End If
       End With
       
        Set tdgBedsReserved.DataSource = rsChargeFileTmp
        
    End If
End Sub

Private Sub Create_ItemsAvailedTmp()

    Set rsChargeFileTmp = Nothing
    Set rsChargeFileTmp = New ADODB.Recordset
    
    With rsChargeFileTmp
        .Fields.Append "chargecode", adInteger
        .Fields.Append "chargename", adVarChar, 70
        .Fields.Append "quantity", adDouble
        .Fields.Append "chargeprice", adDouble
        .Fields.Append "gross_amount", adDouble
        .Fields.Append "discount", adDouble
        .Fields.Append "net_amount", adInteger
        .Open
    End With
    
    With tdgBedsReserved
        Set .DataSource = rsChargeFileTmp
        .ReBind
        .Refresh
        .ReOpen
    End With
    
End Sub

Private Sub tdgregistration_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'lblAll.Caption = ""
    If LastRow <> tdgRegistration.Row + 1 Then
        lblIndividual.Caption = ""
        Create_ItemsAvailedTmp
    End If
    
    If rsRegistration.RecordCount > 0 Then
        cmdCancelReservation.Enabled = True
        cmdCancelBedsReserved.Enabled = True
        If CDate(Format(rsRegistration!arrivaldate, "MM/DD/YYYY")) > CDate(Format(Now, "MM/DD/YYYY")) Then
            cmdRegister.Enabled = False
            cmdSave.Enabled = False
        Else
            cmdRegister.Enabled = True
            cmdSave.Enabled = True
        End If
    Else
        cmdCancelBedsReserved.Enabled = False
        cmdCancelReservation.Enabled = False
        cmdRegister.Enabled = False
        cmdSave.Enabled = False
    End If
    
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchRecord KeyAscii, txtSearch, rsRegistration, txtSearch.Text, tdbSearch.BoundText
  End If
End Sub



