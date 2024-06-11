VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLOBLeave 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9030
   ClientLeft      =   3285
   ClientTop       =   2940
   ClientWidth     =   10860
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
   Icon            =   "frmLeave.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   10860
   Tag             =   "Leave Applications"
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
      ScaleWidth      =   10860
      TabIndex        =   0
      Top             =   0
      Width           =   10860
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves"
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
         TabIndex        =   1
         Top             =   225
         Width           =   5445
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   7890
      Left            =   2010
      TabIndex        =   2
      Top             =   840
      Width           =   8190
      Begin VB.Frame fraButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1170
         TabIndex        =   9
         Top             =   7320
         Width           =   5940
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&NEW"
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
            Image           =   "frmLeave.frx":6852
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   1
            Left            =   1515
            TabIndex        =   11
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&EDIT"
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
            Image           =   "frmLeave.frx":852C
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   8
            Left            =   2970
            TabIndex        =   12
            Top             =   510
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&DELETE"
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
            Image           =   "frmLeave.frx":A206
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   2
            Left            =   2970
            TabIndex        =   13
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
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
            Image           =   "frmLeave.frx":BEE0
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   3
            Left            =   4425
            TabIndex        =   14
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
            Image           =   "frmLeave.frx":CBBA
            cBack           =   14737632
         End
      End
      Begin TrueOleDBGrid80.TDBGrid tdgLeaveLimit 
         Height          =   1950
         Left            =   150
         TabIndex        =   3
         Top             =   1245
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   3440
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Leave"
         Columns(0).DataField=   "leavetypesname"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Limit"
         Columns(1).DataField=   "lvlimit"
         Columns(1).NumberFormat=   "#,##0.00"
         Columns(1).ExternalEditor=   "txtAmount"
         Columns(1).ExternalEditor.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5239"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5159"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3254"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerStyle=0"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3201"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8706"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(1)._HeadDivider=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=79"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         EditDropDown    =   0   'False
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin TDBText6Ctl.TDBText txtFullname 
         Height          =   300
         Left            =   1635
         TabIndex        =   4
         Top             =   570
         Width           =   6030
         _Version        =   65536
         _ExtentX        =   10636
         _ExtentY        =   529
         Caption         =   "frmLeave.frx":D494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave.frx":D500
         Key             =   "frmLeave.frx":D51E
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
         Left            =   7695
         TabIndex        =   5
         ToolTipText     =   "Browse for checked in guests."
         Top             =   570
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
      Begin TrueOleDBGrid80.TDBGrid tdgLeaves 
         Height          =   3600
         Left            =   150
         TabIndex        =   15
         Top             =   3705
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6350
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Date Filed"
         Columns(0).DataField=   "datefiled"
         Columns(0).NumberFormat=   "MM-DD-YY"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Reason"
         Columns(1).DataField=   "remarks"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2672"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2593"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7011"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerStyle=0"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6959"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(1)._HeadDivider=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin TDBNumber6Ctl.TDBNumber txtPayYear 
         Height          =   315
         Left            =   1635
         TabIndex        =   16
         Top             =   195
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmLeave.frx":D562
         Caption         =   "frmLeave.frx":D582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave.frx":D5E8
         Keys            =   "frmLeave.frx":D606
         Spin            =   "frmLeave.frx":D650
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0000"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   4210752
         Format          =   "0000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1996488709
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin lvButton.lvButtons_H cmdUpdateLimit 
         Height          =   465
         Left            =   2865
         TabIndex        =   18
         ToolTipText     =   "Browse for checked in guests."
         Top             =   3210
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   820
         Caption         =   "Update Limit"
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
         Image           =   "frmLeave.frx":D678
         cBack           =   14737632
      End
      Begin TDBNumber6Ctl.TDBNumber txtAmount 
         Height          =   300
         Left            =   375
         TabIndex        =   19
         Top             =   3225
         Visible         =   0   'False
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   529
         Calculator      =   "frmLeave.frx":DDF2
         Caption         =   "frmLeave.frx":DE12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave.frx":DE78
         Keys            =   "frmLeave.frx":DE96
         Spin            =   "frmLeave.frx":DEE0
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#,###,##0.00"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#,###,##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   124518401
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin lvButton.lvButtons_H cmdEditLimit 
         Height          =   465
         Left            =   6585
         TabIndex        =   20
         ToolTipText     =   "Browse for checked in guests."
         Top             =   3225
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   820
         Caption         =   "Edit Limit"
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
         Image           =   "frmLeave.frx":DF08
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdInsertSILApplications 
         Height          =   465
         Left            =   4485
         TabIndex        =   21
         ToolTipText     =   "Browse for checked in guests."
         Top             =   3210
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   820
         Caption         =   "Insert SIL"
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
         Image           =   "frmLeave.frx":E682
         cBack           =   14737632
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   180
         TabIndex        =   17
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Leaves Availed"
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
         Left            =   165
         TabIndex        =   8
         Top             =   3480
         Width           =   3675
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Limit"
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
         Left            =   180
         TabIndex        =   7
         Top             =   1005
         Width           =   3675
      End
      Begin VB.Label Label5 
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
         Left            =   180
         TabIndex        =   6
         Top             =   615
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmLOBLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mEmployeeCode        As Integer
Public mLvHrParam_ID        As Integer
Public mBranchCode          As String
Public mDivisionCode        As String
Public mCostCenterCode      As String

Public rsLeaveLimit         As ADODB.Recordset

Public rsLeaves             As ADODB.Recordset

Private Sub cmdEditLimit_Click()
  If Not IsNull(rsLeaveLimit) Then
    With rsLeaveLimit
      If .RecordCount > 0 Then
        If Not .EOF Then
          
          frmLOBLeave3.mYear = txtPayYear.Text
          frmLOBLeave3.mEmpno = mEmployeeCode
          frmLOBLeave3.mLvCode = rsLeaveLimit!leavetypescode
          frmLOBLeave3.mLvName = rsLeaveLimit!leavetypesname
          frmLOBLeave3.mLvLimit = rsLeaveLimit!lvlimit
          frmLOBLeave3.Show vbModal
          
        End If
      End If
    End With
  End If
End Sub

Private Sub cmdInsertSILApplications_Click()

  cmdInsertSILApplications.Enabled = False
  
  Dim rsTmp         As New ADODB.Recordset
  Dim rsID          As New ADODB.Recordset
  Dim rsLvHrParam   As New ADODB.Recordset
  Dim dblDay        As Double
  Dim dblStart      As Double
  Dim dblEnd        As Double
    
  NetOpen rsTmp, "SELECT x1.employeecode,x2.costcentercode,x2.divisioncode,x2.branchcode,1 leavetypescode, " & _
                  "DATE('2020-03-26') datefiled, DATE('2020-03-26') fromdate, " & _
                  "DATE_ADD('2020-03-26',INTERVAL x1.approved_credit-1 DAY) todate, 'ECQ' remarks, " & _
                  "x1.approved_credit, x2.lvhrparam_id " & _
                  "FROM sil_upload_ver_2 x1 " & _
                  "LEFT OUTER JOIN employee x2 ON x1.employeecode=x2.employeecode"
  
  If rsTmp.RecordCount > 0 Then
    rsTmp.MoveFirst
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    Do While Not rsTmp.EOF
    
      ConMain.Execute "insert into leaveapp_headers (employeecode,costcentercode,divisioncode,branchcode,leavetypescode," & _
                      "datefiled,fromdate,todate,remarks,canceled,trnxdatetime,updtrnxdatetime) values (" & _
                       rsTmp!employeecode & ",'" & rsTmp!costcentercode & "','" & rsTmp!divisioncode & "','" & rsTmp!branchcode & "','" & rsTmp!leavetypescode & "', " & _
                       "'" & Format(rsTmp!datefiled, "YYYY-MM-DD") & "','" & Format(rsTmp!fromdate, "YYYY-MM-DD") & "','" & Format(rsTmp!todate, "YYYY-MM-DD") & "','" & Swap(rsTmp!remarks) & "','N',now(),now())"
      
      NetOpen rsID, "select LAST_INSERT_ID() as last_ID"
      
      NetOpen rsLvHrParam, "select * from lvhr_parameters where lvhrparam_id=" & rsTmp!lvhrparam_id
      
      
      dblStart = CDate(rsTmp!fromdate)
      dblEnd = CDate(rsTmp!todate)
      
'        For dblDay = dblStart To dblEnd Step 1
'            ConMain.Execute "insert into leaveapp_lines (leaveapp_id,employeecode,leavetypescode,leaveapp_date,leaveapp_hours," & _
'                            "firstshift,secondshift,withpay) values (" & rsID!last_ID & "," & rsTmp!employeecode & "," & rsTmp!leavetypescode & ",'" & Format(CDate(dblDay), "YYYY-MM-DD") & "'," & rsLvHrParam(Weekday(CDate(dblDay))).Value & "," & _
'                            "1,1,1)"
'        Next dblDay
      
      For dblDay = dblStart To dblEnd Step 1
        If CDbl(rsLvHrParam(Weekday(CDate(dblDay))).Value) = 0 Then
          dblEnd = dblEnd + 1
        End If
      Next dblDay
      
      ConMain.Execute "update leaveapp_headers set todate = '" & Format(CDate(dblEnd), "YYYY-MM-DD") & "' " & _
                      "where leaveapp_id = " & rsID!last_ID & ";"
      
      For dblDay = dblStart To dblEnd Step 1
'        rsTmpLeaveEntry.AddNew
'        rsTmpLeaveEntry.Fields("leaveapp_date") = CDate(dblDay)
'        rsTmpLeaveEntry.Fields("leaveapp_day") = Format(CDate(dblDay), "dddd")
'        rsTmpLeaveEntry.Fields("leaveapp_hours") = rsLvHrParam(Weekday(CDate(dblDay))).Value
        If CDbl(rsLvHrParam(Weekday(CDate(dblDay))).Value) > 0 Then
'            rsTmpLeaveEntry.Fields("withpay") = 1
'            rsTmpLeaveEntry.Fields("firstshift") = 1
'            rsTmpLeaveEntry.Fields("secondshift") = 1
          ConMain.Execute "insert into leaveapp_lines (leaveapp_id,employeecode,leavetypescode,leaveapp_date,leaveapp_hours," & _
                          "firstshift,secondshift,withpay) values (" & rsID!last_ID & "," & rsTmp!employeecode & "," & rsTmp!leavetypescode & ",'" & Format(CDate(dblDay), "YYYY-MM-DD") & "'," & rsLvHrParam(Weekday(CDate(dblDay))).Value & "," & _
                          "1,1,1)"
        
        End If

'        rsTmpLeaveEntry.Update
      Next dblDay

      rsTmp.MoveNext
    Loop
    ConMain.CommitTrans
    cmdInsertSILApplications.Enabled = False
  End If
End Sub

Private Sub cmdmenu_Click(Index As Integer)
    
    Select Case Index
        
        Case 0:
                frmLOBLeave2.mAdd = True
                frmLOBLeave2.Show vbModal
        Case 1:
                frmLOBLeave2.mAdd = False
                frmLOBLeave2.Show vbModal
        Case 2: Cancel_Clicked
        Case 3: Unload Me
    
    End Select
    
End Sub

Private Sub cmdSearchEmployee_Click()
    With frmBrowseEmployee
        .mBrowseType = "Leaves"
        .mYear = txtPayYear.Text
        .Show vbModal
    End With
End Sub

Private Sub cmdUpdateLimit_Click()
  
  Dim rsTmp         As New ADODB.Recordset
  Dim rsID          As New ADODB.Recordset
  Dim rsLvHrParam   As New ADODB.Recordset
  Dim dblDay        As Double
  Dim dblStart      As Double
  Dim dblEnd        As Double
    
  NetOpen rsTmp, "SELECT x2.employeecode,x2.costcentercode,x2.divisioncode,x2.branchcode, " & _
                 "x1.lvnum,x1.leavetypescode,x1.fromdate,x1.todate,x2.datefiled,x2.remarks,x3.lvhrparam_id FROM lvlne x1 " & _
                 "LEFT OUTER JOIN lvhdr x2 ON x1.lvnum=x2.lvnum " & _
                 "LEFT OUTER JOIN employee x3 on x2.employeecode=x3.employeecode " & _
                 "WHERE x1.todate >= '2017-01-01' " & _
                 "ORDER BY x1.lvnum,x1.fromdate"
  
  If rsTmp.RecordCount > 0 Then
    cmdUpdateLimit.Enabled = True
    rsTmp.MoveFirst
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    Do While Not rsTmp.EOF
      ConMain.Execute "insert into leaveapp_headers (employeecode,costcentercode,divisioncode,branchcode,leavetypescode," & _
                      "datefiled,fromdate,todate,remarks,canceled,trnxdatetime,updtrnxdatetime) values (" & _
                       rsTmp!employeecode & ",'" & rsTmp!costcentercode & "','" & rsTmp!divisioncode & "','" & rsTmp!branchcode & "','" & rsTmp!leavetypescode & "', " & _
                       "'" & Format(rsTmp!datefiled, "YYYY-MM-DD") & "','" & Format(rsTmp!fromdate, "YYYY-MM-DD") & "','" & Format(rsTmp!todate, "YYYY-MM-DD") & "','" & Swap(rsTmp!remarks) & "','N',now(),now())"
        NetOpen rsID, "select LAST_INSERT_ID() as last_ID"
        NetOpen rsLvHrParam, "select * from lvhr_parameters where lvhrparam_id=" & rsTmp!lvhrparam_id
        If CDate(rsTmp!todate) >= CDate("05/26/2017") Then
          dblStart = CDate(rsTmp!fromdate)
          dblEnd = CDate(rsTmp!todate)
          For dblDay = dblStart To dblEnd Step 1
              ConMain.Execute "insert into leaveapp_lines (leaveapp_id,employeecode,leavetypescode,leaveapp_date,leaveapp_hours," & _
                              "firstshift,secondshift,withpay) values (" & rsID!last_ID & "," & rsTmp!employeecode & "," & rsTmp!leavetypescode & ",'" & Format(CDate(dblDay), "YYYY-MM-DD") & "'," & rsLvHrParam(Weekday(CDate(dblDay))).Value & "," & _
                              "1,1,1)"
          Next dblDay
        End If
        rsTmp.MoveNext
    Loop
    ConMain.CommitTrans
    cmdUpdateLimit.Enabled = False
  End If
    
'    If mEmployeeCode <> 0 Then
'        If Not rsLeaveLimit.EOF Then
'            ConMain.Execute "set autocommit = 0"
'            ConMain.BeginTrans
'            ConMain.Execute "delete from lvlimit where payyear = " & txtPayYear.Text & " and employeecode = " & mEmployeeCode & ""
'
'            Do While Not rsLeaveLimit.EOF
'                ConMain.Execute "insert into lvlimit(payyear,employeecode,leavetypescode,lvlimit) values (" & _
'                                                    txtPayYear.Text & ", " & mEmployeeCode & "," & rsLeaveLimit!leavetypescode & "," & rsLeaveLimit!lvlimit & ") "
'                rsLeaveLimit.MoveNext
'            Loop
'            ConMain.CommitTrans
'
'            MsgBox "Leave limits were succesfully udpated.", vbInformation + vbOKOnly
'        End If
'    Else
'        MsgBox "Please select and employee.", vbExclamation + vbOKOnly
'    End If
End Sub

Private Sub Form_Activate()

    Focus_MDIButton Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Remove_MDIButton Me.Name
    
End Sub

Private Sub Form_Load()
    Add_MDIButton Me.Name, Me.Tag
    Lock_Button "FFFT", frmLOBLeave.cmdMenu, 3
    txtPayYear.Text = Format(Now, "YYYY")
    Create_TmpLeaveLimit
End Sub

Private Sub Form_Resize()
    
    With fraMain
        .Top = ((Me.ScaleHeight / 2) + 300) - ((.Height / 2))
        .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    End With
    
End Sub

Private Sub Cancel_Clicked()
        
    With rsLeaves
        If .RecordCount > 0 Then
            If Not .EOF Then
                If MsgBox("Do you want to cancel this leave entry?", vbQuestion + vbYesNo) = vbYes Then
                    ConMain.Execute "set autocommit = 0"
                    ConMain.BeginTrans
                    ConMain.Execute "update leaveapp_headers set cancel = 'Y' where leaveapp_id = " & !leaveapp_id & ""
                    ConMain.Execute "delete from leaveapp_lines where leaveapp_id = " & !leaveapp_id & " "
                    ConMain.CommitTrans
                    rsLeaves.Requery
                    If rsLeaves.RecordCount > 0 Then
                      Lock_Button "TTTT", cmdMenu, 3
                    Else
                      Lock_Button "TFFT", cmdMenu, 3
                    End If
                End If
            End If
        End If
    End With
    
End Sub

Public Sub Create_TmpLeaveLimit()

  Set rsLeaveLimit = Nothing
  Set rsLeaveLimit = New ADODB.Recordset
  
  With rsLeaveLimit
    .Fields.Append "leavetypescode", adVarChar, 7
    .Fields.Append "leavetypesname", adVarChar, 50
    .Fields.Append "lvlimit", adDouble
    .Open
  End With
  
  Set tdgLeaveLimit.DataSource = rsLeaveLimit

End Sub

Private Sub txtAmount_LostFocus()
    On Error Resume Next
    
    tdgLeaveLimit.SetFocus
End Sub

Private Sub txtPayYear_LostFocus()
  
  Dim rsTmp       As ADODB.Recordset
  
  If mEmployeeCode > 0 Then
  
'       NetOpen rsTmp, "select x1.*,concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname, " & _
'           "x3.costcenter,x4.division,x5.branch " & _
'           "from (select * from lvhdr where employeecode =" & mEmployeeCode & " and (year(tdate) = '" & txtPayYear.Text & "' or year(datefiled) = '" & txtPayYear.Text & "'))  x1 " & _
'           "left outer join employee x2 on x1.employeecode = x2.employeecode " & _
'           "left outer join costcenter x3 on x1.costcentercode = x3.costcentercode " & _
'           "left outer join division x4 on x1.divisioncode = x4.divisioncode " & _
'           "left outer join Branch x5 on x1.branchcode = x5.branchcode " & _
'           "where x1.cancel = 'N'  " & _
'           "order by x1.datefiled desc"
           '"order by concat(x2.lastname,', ',x2.firstname,' ',x2.middlename)"
    
        NetOpen rsTmp, "select x1.*,concat(x2.lastname,', ',x2.firstname,' ',x2.middlename) fullname, " & _
            "x3.costcenter,x4.division,x5.branch,x6.leavetypesname " & _
            "from (select * from leaveapp_headers where employeecode =" & mEmployeeCode & " and year(trnxdatetime) = '" & txtPayYear.Text & "')  x1 " & _
            "left outer join employee x2 on x1.employeecode = x2.employeecode " & _
            "left outer join costcenter x3 on x1.costcentercode = x3.costcentercode " & _
            "left outer join division x4 on x1.divisioncode = x4.divisioncode " & _
            "left outer join Branch x5 on x1.branchcode = x5.branchcode " & _
            "left outer join leavetypes x6 on x1.leavetypescode=x6.leavetypescode " & _
            "where x1.canceled = 'N'  " & _
            "order by x1.leaveapp_id desc"
                        
       Set rsLeaves = rsTmp.Clone
       Set tdgLeaves.DataSource = rsLeaves
       
  End If
  
End Sub
