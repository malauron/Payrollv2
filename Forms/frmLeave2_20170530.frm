VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLOBLeave2_20170530 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Application Details"
   ClientHeight    =   4800
   ClientLeft      =   3390
   ClientTop       =   3390
   ClientWidth     =   13080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeave2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   30
      TabIndex        =   11
      Top             =   4305
      Width           =   13035
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   15
         TabIndex        =   6
         Top             =   60
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
         Image           =   "frmLeave2.frx":6852
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   390
         Left            =   2010
         TabIndex        =   5
         Top             =   60
         Width           =   1995
         _ExtentX        =   3519
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
         Image           =   "frmLeave2.frx":752C
         cBack           =   14737632
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4365
      Left            =   30
      TabIndex        =   8
      Top             =   -75
      Width           =   13050
      Begin TDBText6Ctl.TDBText txtLvNum 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   210
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmLeave2.frx":8206
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":8272
         Key             =   "frmLeave2.frx":8290
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
      Begin TDBText6Ctl.TDBText txtRemarks 
         Height          =   960
         Left            =   6480
         TabIndex        =   3
         Top             =   210
         Width           =   6510
         _Version        =   65536
         _ExtentX        =   11483
         _ExtentY        =   1693
         Caption         =   "frmLeave2.frx":82D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":8340
         Key             =   "frmLeave2.frx":835E
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
         AlignVertical   =   0
         MultiLine       =   -1
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
      Begin TrueOleDBGrid80.TDBGrid tdgLeaveEntry 
         Height          =   3060
         Left            =   1560
         TabIndex        =   4
         ToolTipText     =   "Note: Press the DELETE button to remove a leave entry."
         Top             =   1245
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   5398
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Code"
         Columns(0).DataField=   "leavetypescode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Leave Type"
         Columns(1).DataField=   "leavetypesname"
         Columns(1).DropDown=   "tddLvList"
         Columns(1).DropDown.vt=   8
         Columns(1).ExternalEditor=   "txtleavetypes"
         Columns(1).ExternalEditor.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Start Date"
         Columns(2).DataField=   "fromdate"
         Columns(2).NumberFormat=   "MM-DD-YY"
         Columns(2).ExternalEditor=   "txtlvDate"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "End Date"
         Columns(3).DataField=   "todate"
         Columns(3).NumberFormat=   "MM-DD-YY"
         Columns(3).ExternalEditor=   "txtlvDate"
         Columns(3).ExternalEditor.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "First Shift"
         Columns(4).DataField=   "firstshift"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Second Shift"
         Columns(5).DataField=   "secondshift"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "With Pay"
         Columns(6).DataField=   "withpay"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8708"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=7514"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7435"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1746"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=1720"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1640"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=2117"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2037"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=2090"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerStyle=0"
         Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2037"
         Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(45)=   "Column(6)._HeadDivider=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Width=79"
         Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=8708"
         Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
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
      Begin TrueOleDBGrid80.TDBDropDown tddLvList 
         Height          =   1365
         Left            =   75
         TabIndex        =   16
         Top             =   3525
         Visible         =   0   'False
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   2408
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Code"
         Columns(0).DataField=   "leavetypescode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Name"
         Columns(1).DataField=   "leavetypesname"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "withpay"
         Columns(2).DataField=   "withpay"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).AllowColMove=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         AllowRowSizing  =   -1  'True
         Appearance      =   2
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   0
         RowDividerStyle =   0
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   -1  'True
         ListField       =   "Branch"
         DataField       =   ""
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   -1  'True
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   14215660
         ValueTranslate  =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFFF0EA&"
         _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
         _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFFFF&"
         _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
         _StyleDefs(17)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(18)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(19)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(20)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(21)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(22)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(23)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(24)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(25)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(26)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(27)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(28)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(29)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(30)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(39)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Named:id=33:Normal"
         _StyleDefs(44)  =   ":id=33,.parent=0"
         _StyleDefs(45)  =   "Named:id=34:Heading"
         _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   ":id=34,.wraptext=-1"
         _StyleDefs(48)  =   "Named:id=35:Footing"
         _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   "Named:id=36:Selected"
         _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(52)  =   "Named:id=37:Caption"
         _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(54)  =   "Named:id=38:HighlightRow"
         _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=39:EvenRow"
         _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(58)  =   "Named:id=40:OddRow"
         _StyleDefs(59)  =   ":id=40,.parent=33"
         _StyleDefs(60)  =   "Named:id=41:RecordSelector"
         _StyleDefs(61)  =   ":id=41,.parent=34"
         _StyleDefs(62)  =   "Named:id=42:FilterBar"
         _StyleDefs(63)  =   ":id=42,.parent=33"
      End
      Begin TDBText6Ctl.TDBText txtleavetypes 
         Height          =   300
         Left            =   105
         TabIndex        =   15
         Top             =   3255
         Visible         =   0   'False
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmLeave2.frx":83A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":840E
         Key             =   "frmLeave2.frx":842C
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
         BorderStyle     =   0
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
      Begin TDBDate6Ctl.TDBDate txtlvDate 
         Height          =   300
         Left            =   1110
         TabIndex        =   14
         Top             =   3045
         Visible         =   0   'False
         Width           =   1650
         _Version        =   65536
         _ExtentX        =   2910
         _ExtentY        =   529
         Calendar        =   "frmLeave2.frx":8470
         Caption         =   "frmLeave2.frx":8576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":85DC
         Keys            =   "frmLeave2.frx":85FA
         Spin            =   "frmLeave2.frx":8658
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
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
      Begin TDBDate6Ctl.TDBDate txtDateFiled 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   900
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Calendar        =   "frmLeave2.frx":8680
         Caption         =   "frmLeave2.frx":8786
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":87EC
         Keys            =   "frmLeave2.frx":880A
         Spin            =   "frmLeave2.frx":8868
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
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
      Begin TDBText6Ctl.TDBText txtFullname 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   555
         Width           =   4845
         _Version        =   65536
         _ExtentX        =   8546
         _ExtentY        =   529
         Caption         =   "frmLeave2.frx":8890
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":88FC
         Key             =   "frmLeave2.frx":891A
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -105
         TabIndex        =   17
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   4425
         TabIndex        =   13
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Entries:"
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
         Left            =   285
         TabIndex        =   12
         Top             =   2775
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Filed"
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
         Height          =   255
         Left            =   -120
         TabIndex        =   10
         Top             =   945
         Width           =   1560
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Number"
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
         Height          =   255
         Left            =   -105
         TabIndex        =   9
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "modiifie"
      Default         =   -1  'True
      Height          =   375
      Left            =   13185
      TabIndex        =   7
      Top             =   3105
      Width           =   1215
   End
End
Attribute VB_Name = "frmLOBLeave2_20170530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd                 As Boolean
Dim rsTmpLeaveEntry         As ADODB.Recordset

Private Sub Form_Load()

  Dim rsLeaveEntry  As ADODB.Recordset

  Create_TmpLeaveEntry

'  bind_tdb ConMain, tdbEmpNo, "select employeecode,dummycode from employee order by dummycode", "dummycode", "employeecode"
'
'  bind_tdb ConMain, tdbEmpName, "select x1.employeecode,concat(x1.lastname,', ',x1.firstname,' ',x1.middlename) fullname, " & _
'                          "x1.costcentercode, x1.divisioncode, x1.branchcode,x2.costcenter,x3.division,x4.branch from employee x1 " & _
'                          "left outer join costcenter x2 on x1.costcentercode = x2.costcentercode " & _
'                          "left outer join division x3 on x1.divisioncode = x3.divisioncode " & _
'                          "left outer join branch x4 on x1.branchcode = x4.branchcode order by x1.lastname,x1.firstname,x1.middlename", "fullname", "employeecode"

  Bind_tdd ConMain, tddLvList, "select * from leavetypes order by leavetypesname", "leavetypesname"

    txtFullname.Text = frmLOBLeave.txtFullname
    
  If mAdd = False Then
    With frmLOBLeave.rsLeaves
      If .RecordCount > 0 Then

        NetOpen rsLeaveEntry, "select x1.*,x2.leavetypesname from lvlne x1 " & _
                                  "left outer join leavetypes x2 on x1.leavetypescode = x2.leavetypescode " & _
                                  "where x1.lvnum = '" & !lvnum & "'"

        txtLvNum.Text = !lvnum
        txtDateFiled.Text = Format(!datefiled, "MM/DD/YYYY")
        
        txtRemarks.Text = !remarks

        If rsLeaveEntry.RecordCount > 0 Then
          rsLeaveEntry.MoveFirst
          Do While Not rsLeaveEntry.EOF
            rsTmpLeaveEntry.AddNew
            rsTmpLeaveEntry.Fields("leavetypescode") = rsLeaveEntry!leavetypescode
            rsTmpLeaveEntry.Fields("leavetypesname") = rsLeaveEntry!LeaveTypesname
            rsTmpLeaveEntry.Fields("fromdate") = rsLeaveEntry!fromdate
            rsTmpLeaveEntry.Fields("todate") = rsLeaveEntry!todate
            rsTmpLeaveEntry.Fields("withpay") = rsLeaveEntry!withpay
            rsTmpLeaveEntry.Fields("firstshift") = rsLeaveEntry!firstshift
            rsTmpLeaveEntry.Fields("secondshift") = rsLeaveEntry!secondshift
            rsTmpLeaveEntry.Update
            rsLeaveEntry.MoveNext
          Loop
        End If
      End If
    End With
  Else
    txtDateFiled.Text = ""
  End If

End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorHandler

    Dim rsChk               As ADODB.Recordset
    Dim mAppDate            As Date

    Dim mErrorType          As String

    If Not IsDate(txtDateFiled.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtDateFiled.SetFocus
      Exit Sub
    End If

'    If Trim(tdbEmpName.Text) = "" Or IsNull(tdbEmpName.SelectedItem) Or tdbEmpName.ApproxCount = 0 Then
'      MsgBox "Please select an employee.", vbExclamation + vbOKOnly
'      tdbEmpName.SetFocus
'      Exit Sub
'    End If
'
'    If rsTmpLeaveEntry.RecordCount <= 0 Then
'      MsgBox "Please enter at least one (1) leave entry.", vbExclamation + vbOKOnly
'      tdgLeaveEntry.SetFocus
'      Exit Sub
'    End If

    With rsTmpLeaveEntry
      .MoveFirst
      Do While Not .EOF
        If !leavetypescode <> "" Then

            If Not IsDate(!fromdate) Then
                MsgBox "Please enter a valid date."
                tdgLeaveEntry.SetFocus
                Exit Sub
            End If

            If Not IsDate(!todate) Then
                MsgBox "Please enter a valid date."
                tdgLeaveEntry.SetFocus
                Exit Sub
            End If

            If CDate(Format(!fromdate, "MM/DD/YYYY")) > CDate(Format(!todate, "MM/DD/YYYY")) Then
                MsgBox "End date is earlier than the start date.", vbExclamation + vbOKOnly
                tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
                tdgLeaveEntry.SetFocus
                Exit Sub
            End If

            If !firstshift = 0 And !secondshift = 0 Then
                MsgBox "You must select at least one shift.", vbExclamation + vbOKOnly
                tdgLeaveEntry.Col = tdgLeaveEntry.Columns("firstshift").ColIndex
                tdgLeaveEntry.SetFocus
                Exit Sub
            End If

        End If
        .MoveNext
        
      Loop
    End With

    If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If

    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    With frmLOBLeave
        If mAdd = True Then
          txtLvNum.Text = LastCode("Leaves")
          ConMain.Execute "insert into lvhdr(lvnum,employeecode,costcentercode, " & _
                                "divisioncode,branchcode,datefiled,tdate,ttime,remarks,cancel) values ('" & txtLvNum.Text & "', " & _
                                 .mEmployeeCode & "," & .mCostCenterCode & "," & .mDivisionCode & ", " & _
                                .mBranchCode & ",'" & Format(txtDateFiled.Text, "YYYY-MM-DD") & "', " & _
                                "'" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "hh:nn") & "','" & Swap(txtRemarks.Text) & "','N')"
        Else
          ConMain.Execute "update lvhdr set employeecode = " & .mEmployeeCode & ", costcentercode = " & .mCostCenterCode & ", " & _
                                "divisioncode = " & .mDivisionCode & ", branchcode = " & .mBranchCode & ", " & _
                                "datefiled = '" & Format(txtDateFiled.Text, "YYYY-MM-DD") & "', tdate = '" & Format(Now, "YYYY-MM-DD") & "', " & _
                                "ttime = '" & Format(Now, "hh:nn") & "',remarks = '" & Swap(txtRemarks.Text) & "' where lvnum = '" & txtLvNum.Text & "'"
          ConMain.Execute "delete from lvlne where lvnum = '" & txtLvNum.Text & "'"
          ConMain.Execute "delete from appdate where apptype = 'LV' and trnxcode = '" & txtLvNum.Text & "'"
        End If
    End With
    
    With rsTmpLeaveEntry
        .MoveFirst
        Do While Not .EOF
            ConMain.Execute "insert into lvlne(lvnum,leavetypescode,fromdate,todate,firstshift,secondshift,withpay) values " & _
                            "('" & txtLvNum.Text & "','" & !leavetypescode & "', " & _
                            "'" & Format(!fromdate, "YYYY-MM-DD") & "','" & Format(!todate, "YYYY-MM-DD") & "','" & IIf(!firstshift <= 0, 0, 1) & "','" & IIf(!secondshift <= 0, 0, 1) & "','" & IIf(!withpay <> 0, 1, 0) & "')"
            mAppDate = Format(!fromdate, "MM/DD/YYYY")
            Do While mAppDate <= CDate(Format(!todate, "MM/DD/YYYY"))
                mErrorType = "Duplicate entry"
                If !firstshift = 1 And !secondshift = 1 Then
                    NetOpen rsChk, "select * from appdate where trnxdate = '" & Format(mAppDate, "YYYY-MM-DD") & "' and employeecode = " & frmLOBLeave.mEmployeeCode & " and (firstshift = 1 or secondshift = 1)"
                    If rsChk.RecordCount > 0 Then
                        If rsChk!apptype = "LV" Then
                            MsgBox "Employee has already a leave application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
                            ConMain.RollbackTrans
                            tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
                            tdgLeaveEntry.SetFocus
                            Exit Sub
                        ElseIf rsChk!apptype = "OBT" Then
                            MsgBox "Employee has a already travel application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
                            ConMain.RollbackTrans
                            tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
                            tdgLeaveEntry.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
                ConMain.Execute "insert into appdate (apptype,employeecode,trnxcode,trnxdate,leavetypescode,firstshift,secondshift,withpay,processed) values ('LV'," & frmLOBLeave.mEmployeeCode & ",'" & txtLvNum.Text & "','" & Format(mAppDate, "YYYY-MM-DD") & "'," & !leavetypescode & ",'" & IIf(!firstshift <= 0, 0, 1) & "','" & IIf(!secondshift <= 0, 0, 1) & "','" & IIf(!withpay <> 0, 1, 0) & "','N')"
                mAppDate = mAppDate + 1
            Loop
            .MoveNext
        Loop
    End With

    ConMain.CommitTrans

    frmLOBLeave.rsLeaves.Requery
    frmLOBLeave.rsLeaves.Find "lvnum = '" & txtLvNum.Text & "'"
    Lock_Button "TTTT", frmLOBLeave.cmdMenu, 3

    Unload Me

    Exit Sub

ErrorHandler:

    If mErrorType = "Duplicate entry" Then

        With rsTmpLeaveEntry
            NetOpen rsChk, "select * from appdate where trnxdate = '" & Format(mAppDate, "YYYY-MM-DD") & "' and employeecode = " & frmLOBLeave.mEmployeeCode & " and firstshift = " & IIf(!firstshift <= 0, 0, 1) & " and secondshift = " & IIf(!secondshift <= 0, 0, 1) & ""
            If rsChk.RecordCount > 0 Then
                If rsChk!apptype = "LV" Then
                    MsgBox "Employee has already a leave application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
                ElseIf rsChk!apptype = "OBT" Then
                    MsgBox "Employee has a already travel application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
                End If
            End If
        End With

        ConMain.RollbackTrans

    End If

End Sub

Private Sub Create_TmpLeaveEntry()

  Set rsTmpLeaveEntry = Nothing
  Set rsTmpLeaveEntry = New ADODB.Recordset

  With rsTmpLeaveEntry
    .Fields.Append "leavetypescode", adVarChar, 7
    .Fields.Append "leavetypesname", adVarChar, 50
    .Fields.Append "fromdate", adDate
    .Fields.Append "todate", adDate
    .Fields.Append "withpay", adInteger
    .Fields.Append "firstshift", adInteger
    .Fields.Append "secondshift", adInteger
    .Open
  End With

  Set tdgLeaveEntry.DataSource = rsTmpLeaveEntry

End Sub

Private Sub tddLvList_DropDownOpen()
    Bind_tdd ConMain, tddLvList, "select * from leavetypes order by leavetypesname", "leavetypesname"
    With tddLvList
        .Width = tdgLeaveEntry.Columns("leavetypesname").Width
    End With
End Sub

Private Sub tddLvList_RowChange()

  With tdgLeaveEntry
    .Columns("leavetypescode").Text = tddLvList.Columns("leavetypescode").Text
    txtleavetypes.Text = tddLvList.Columns("leavetypesname").Text
    .Columns("withpay").Value = tddLvList.Columns("withpay").Text
    .Columns("firstshift").Value = 1
    .Columns("secondshift").Value = 1
    If Not IsDate(.Columns("fromdate").Text) Then
      .Columns("fromdate").Text = Format(Now, "MM/DD/YYYY")
    End If
    If Not IsDate(.Columns("todate").Text) Then
      .Columns("todate").Text = Format(Now, "MM/DD/YYYY")
    End If
  End With

End Sub

Private Sub tdgLeaveEntry_AfterColUpdate(ByVal ColIndex As Integer)
    With tdgLeaveEntry
        If ColIndex = .Columns("fromdate").ColIndex Or ColIndex = .Columns("todate").ColIndex Then
            If Format(.Columns("fromdate").Text, "MM/DD/YYYY") <> Format(.Columns("todate").Text, "MM/DD/YYYY") Then
                .Columns("firstshift").Value = 1
                .Columns("secondshift").Value = 1
            End If
        End If
    End With
End Sub

Private Sub tdgLeaveEntry_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

  With tdgLeaveEntry

      If ColIndex = .Columns("fromdate").ColIndex Or _
         ColIndex = .Columns("todate").ColIndex Or _
         ColIndex = .Columns("firstshift").ColIndex Or _
         ColIndex = .Columns("secondshift").ColIndex Or _
         ColIndex = .Columns("withpay").ColIndex Then

        If .Columns("leavetypescode").Text = "" Then Cancel = True

        If ColIndex = .Columns("firstshift").ColIndex Or ColIndex = .Columns("secondshift").ColIndex Then
            If Format(.Columns("fromdate").Text, "MM/DD/YYYY") <> Format(.Columns("todate").Text, "MM/DD/YYYY") Then
                Cancel = True
            End If
        End If

        If ColIndex = .Columns("firstshift").ColIndex Then
            If .Columns("secondshift").Value = 0 Then
                Cancel = True
            End If
        End If

        If ColIndex = .Columns("secondshift").ColIndex Then
            If .Columns("firstshift").Value = 0 Then
                Cancel = True
            End If
        End If
      End If

    End With

End Sub

Private Sub tdgLeaveEntry_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If ColIndex = tdgLeaveEntry.Columns("fromdate").ColIndex Or ColIndex = tdgLeaveEntry.Columns("todate").ColIndex Then
    If Not IsDate(txtlvDate.Text) Then
        Cancel = True
    End If
  End If
End Sub

Private Sub tdgLeaveEntry_BeforeRowColChange(Cancel As Integer)

  If tdgLeaveEntry.Columns("leavetypescode").Text <> "" Then

    If tdgLeaveEntry.Columns("leavetypesname").Text = "" Then
      Cancel = True
      tdgLeaveEntry.Col = tdgLeaveEntry.Columns("leavetypesname").ColIndex
      SendKeys " "
      Exit Sub
    End If

    If tdgLeaveEntry.Columns("fromdate").Text = "" Then
      Cancel = True
      tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
      Exit Sub
    End If

    If tdgLeaveEntry.Columns("todate").Text = "" Then
      Cancel = True
      tdgLeaveEntry.Col = tdgLeaveEntry.Columns("todate").ColIndex
      Exit Sub
    End If

    If tdgLeaveEntry.Columns("firstshift").Value = 0 And tdgLeaveEntry.Columns("secondshift").Value = 0 Then
      tdgLeaveEntry.Columns("firstshift").Value = 1
      tdgLeaveEntry.Columns("secondshift").Value = 1
    End If

  End If

End Sub

Private Sub tdgLeaveEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If txtleavetypes.Visible = False And txtlvDate.Visible = False Then
            With tdgLeaveEntry
                If rsTmpLeaveEntry.RecordCount > 0 Then
                    If Not .EOF And Not .BOF Then
                        If MsgBox("Do you want to delete this entry?", vbQuestion + vbYesNo) = vbYes Then
                            .Delete
                            .Refresh
                        End If
                        .SetFocus
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub txtDateFiled_GotFocus()
    With txtDateFiled
        .SelStart = 0
        .SelLength = Len(txtDateFiled.Text)
    End With
End Sub

Private Sub txtDateFiled_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFullname_Keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtLeaveTypes_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tdgLeaveEntry.SetFocus
  Else
    SearchRecord KeyAscii, txtleavetypes, tddLvList.DataSource, txtleavetypes.Text, "leavetypesname"
    tddLvList_RowChange
  End If
End Sub

Private Sub txtLeaveTypes_LostFocus()
  tdgLeaveEntry.SetFocus
End Sub

Private Sub txtLvDate_LostFocus()
  tdgLeaveEntry.SetFocus
End Sub

Private Sub txtLvNum_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtRemarks_GotFocus()
    With txtRemarks
        .SelStart = 0
        .SelLength = Len(txtRemarks.Text)
    End With
End Sub

Private Sub txtRemarks_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "TAB"
    End If
End Sub
