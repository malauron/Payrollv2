VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Begin VB.Form frmProvince 
   BackColor       =   &H80000016&
   Caption         =   "Province"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      BackColor       =   &H80000016&
      Height          =   1335
      Left            =   105
      TabIndex        =   13
      Top             =   1350
      Width           =   7125
      Begin TDBText6Ctl.TDBText txtProvcode 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   255
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3528
         _ExtentY        =   529
         Caption         =   "frmProvince.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmProvince.frx":006C
         Key             =   "frmProvince.frx":008A
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
      Begin TDBText6Ctl.TDBText txtDescription 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   915
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   529
         Caption         =   "frmProvince.frx":00CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmProvince.frx":013A
         Key             =   "frmProvince.frx":0158
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
         MaxLength       =   100
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
      Begin TDBText6Ctl.TDBText txtProvname 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   585
         Width           =   5400
         _Version        =   65536
         _ExtentX        =   9525
         _ExtentY        =   529
         Caption         =   "frmProvince.frx":019C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmProvince.frx":0208
         Key             =   "frmProvince.frx":0226
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
         MaxLength       =   50
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
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
         Left            =   -30
         TabIndex        =   16
         Top             =   645
         Width           =   1470
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   -120
         TabIndex        =   15
         Top             =   345
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
         Left            =   210
         TabIndex        =   14
         Top             =   975
         Width           =   1230
      End
   End
   Begin VB.Frame fraButton 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8310
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   5
         Left            =   5850
         TabIndex        =   9
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmProvince.frx":026A
         PICN            =   "frmProvince.frx":0286
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   4
         Left            =   4710
         TabIndex        =   8
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmProvince.frx":0820
         PICN            =   "frmProvince.frx":083C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   2
         Left            =   2430
         TabIndex        =   6
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmProvince.frx":0DD8
         PICN            =   "frmProvince.frx":0DF4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   1
         Left            =   1290
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmProvince.frx":1390
         PICN            =   "frmProvince.frx":13AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   3
         Left            =   3570
         TabIndex        =   7
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmProvince.frx":1948
         PICN            =   "frmProvince.frx":1964
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdProvince 
         Height          =   465
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   820
         BTYPE           =   8
         TX              =   "&New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483626
         BCOLO           =   -2147483626
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmProvince.frx":1F00
         PICN            =   "frmProvince.frx":1F1C
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
   Begin VB.Frame fraSearch 
      BackColor       =   &H80000016&
      Height          =   720
      Left            =   105
      TabIndex        =   10
      Top             =   615
      Width           =   5895
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   300
         Left            =   1395
         TabIndex        =   3
         Top             =   255
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   529
         Caption         =   "frmProvince.frx":24B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmProvince.frx":2524
         Key             =   "frmProvince.frx":2542
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH"
         Height          =   255
         Left            =   375
         TabIndex        =   11
         Top             =   315
         Width           =   915
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgProvince 
      Height          =   2310
      Left            =   150
      TabIndex        =   17
      Top             =   2805
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4075
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "provcode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Province"
      Columns(1).DataField=   "provname"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Description"
      Columns(2).DataField=   "description"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3625"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3545"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1852"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1773"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H400000&"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
      _StyleDefs(16)  =   ":id=8,.fgcolor=&HFFFFFF&"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HD7F9FD&"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(45)  =   "Named:id=33:Normal"
      _StyleDefs(46)  =   ":id=33,.parent=0"
      _StyleDefs(47)  =   "Named:id=34:Heading"
      _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   ":id=34,.wraptext=-1"
      _StyleDefs(50)  =   "Named:id=35:Footing"
      _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(58)  =   "Named:id=39:EvenRow"
      _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=40:OddRow"
      _StyleDefs(61)  =   ":id=40,.parent=33"
      _StyleDefs(62)  =   "Named:id=41:RecordSelector"
      _StyleDefs(63)  =   ":id=41,.parent=34"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmProvince"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProvince             As ADODB.Recordset
Dim mSort                  As String

Private Sub cmdprovince_Click(Index As Integer)
  Select Case Index
    Case 0: AddSave_Button_Clicked
    Case 1: EditUpdate_Button_Clicked
    Case 2:
    Case 3: Cancel_Clicked
    Case 4:
    Case 5: Unload Me
  End Select
End Sub

Private Sub Form_Load()

  NetOpen rsProvince, "", "select * from province order by provcode "
  If rsProvince.RecordCount > 0 Then
    rsProvince.MoveFirst
  End If
  Set tdgProvince.Datasource = rsProvince
  mSort = "provcode"
    
  cmdprovince_Click 3
  
End Sub

Private Sub AddSave_Button_Clicked()

  If cmdProvince(0).Caption = "&New" Then
  
    Lock_Button "TFFTFF", cmdProvince, 5
    cmdProvince(0).Caption = "&Save"
    ClearText
    txtProvcode.Text = "...."
    fraSearch.Enabled = False
    fraInfo.Enabled = True
    tdgProvince.Enabled = False
    txtProvname.SetFocus
  
  Else
  
    If Trim(txtProvname.Text) = "" Then
      MsgBox "Province name is blank.", vbExclamation + vbOKOnly
      txtProvname.SetFocus
      Exit Sub
    End If
    
    CitronPayroll.Execute "set autocommit = 0"
    CitronPayroll.BeginTrans
    txtProvcode.Text = LastCode("GetLastCodeA", "Province", "0000000")
    CitronPayroll.Execute "insert into province (provcode,provname,description) values ('" & txtProvcode.Text & "','" & txtProvname.Text & "','" & txtDescription.Text & "')"
    CitronPayroll.CommitTrans
    rsProvince.Requery
    rsProvince.Find "provcode = '" & txtProvcode.Text & "'"
        
    cmdprovince_Click 3
      
  End If
  
End Sub

Private Sub EditUpdate_Button_Clicked()

  If cmdProvince(1).Caption = "&Edit" Then
    
      Lock_Button "FTFTFF", cmdProvince, 5
      cmdProvince(1).Caption = "&Update"
      fraSearch.Enabled = False
      fraInfo.Enabled = True
      tdgProvince.Enabled = False
      txtProvname.SetFocus
  
  Else
  
  
    If Trim(txtProvname.Text) = "" Then
      MsgBox "Province name is blank.", vbExclamation + vbOKOnly
      txtProvname.SetFocus
      Exit Sub
    End If
    
    CitronPayroll.Execute "update province set provname = '" & txtProvname.Text & "', description = '" & txtDescription.Text & "' where provcode = '" & txtProvcode.Text & "'"
                       
    rsProvince.Requery
    rsProvince.Find "provcode = '" & txtProvcode.Text & "'"
        
    cmdprovince_Click 3
    
  End If
  
End Sub

Private Sub Cancel_Clicked()

  If rsProvince.RecordCount > 0 Then
    Lock_Button "TTTFTT", cmdProvince, 5
  Else
    Lock_Button "TFFFTT", cmdProvince, 5
  End If

  cmdProvince(0).Caption = "&New"
  cmdProvince(1).Caption = "&Edit"

  fraSearch.Enabled = True
  fraInfo.Enabled = False
  tdgProvince.Enabled = True
  tdgprovince_RowColChange 0, 0
  
End Sub


Private Sub Form_Resize()

  With fraButton
    .Top = 0
    .Left = 0
  End With
  
  With fraSearch
    .Top = fraButton.Top + fraButton.Height
    .Left = 150
    .Width = Me.ScaleWidth - 300
  End With
  
  With fraInfo
    .Top = fraSearch.Top + fraSearch.Height
    .Left = 150
    .Width = Me.ScaleWidth - 300
  End With
  
  With tdgProvince
    .Top = fraInfo.Top + fraInfo.Height
    .Left = 150
    .Width = Me.ScaleWidth - 300
    .Height = Me.ScaleHeight - .Top
  End With
  
End Sub

Private Sub tdgprovince_HeadClick(ByVal ColIndex As Integer)

    If rsProvince.RecordCount > 0 Then
      mSort = tdgProvince.Columns(ColIndex).Datafield
      rsProvince.Sort = mSort
    End If
  
End Sub

Private Sub tdgprovince_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

  With rsProvince
    If .RecordCount > 0 Then
      txtProvcode.Text = !provcode
      txtProvname.Text = !provname
      txtDescription.Text = !Description
    Else
      ClearText
    End If
  End With
  
End Sub

Private Sub txtSearch_Keypress(Keyascii As Integer)
  If Keyascii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchRecord Keyascii, txtSearch, rsProvince, txtSearch.Text, mSort
  End If
End Sub

Private Sub ClearText()

    txtProvcode.Text = ""
    txtProvname.Text = ""
    txtDescription.Text = ""
    
End Sub







