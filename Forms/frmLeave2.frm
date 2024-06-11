VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLOBLeave2 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Application Details"
   ClientHeight    =   7230
   ClientLeft      =   3390
   ClientTop       =   3390
   ClientWidth     =   13140
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
   Icon            =   "frmLeave2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   30
      TabIndex        =   13
      Top             =   6750
      Width           =   13035
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   15
         TabIndex        =   8
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
         TabIndex        =   7
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
      Height          =   6750
      Left            =   15
      TabIndex        =   10
      Top             =   -30
      Width           =   13050
      Begin lvButton.lvButtons_H cmdLeaveType 
         Height          =   315
         Left            =   4545
         TabIndex        =   21
         Top             =   870
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
      Begin TrueOleDBGrid80.TDBGrid tdglvDays 
         Height          =   5160
         Left            =   1560
         TabIndex        =   17
         Top             =   1530
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9102
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Date"
         Columns(0).DataField=   "leaveapp_date"
         Columns(0).NumberFormat=   "MM-DD-YY"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Day of the week"
         Columns(1).DataField=   "leaveapp_day"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Total Working Hours"
         Columns(2).DataField=   "leaveapp_hours"
         Columns(2).NumberFormat=   "#,##0"
         Columns(2).ExternalEditor=   "txtAmount"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "First Shift"
         Columns(3).DataField=   "firstshift"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Second Shift"
         Columns(4).DataField=   "secondshift"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "With Pay"
         Columns(5).DataField=   "withpay"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8708"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(3).Width=2117"
         Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2037"
         Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=2117"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2037"
         Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(33)=   "Column(5).Width=2090"
         Splits(0)._ColumnProps(34)=   "Column(5).DividerStyle=0"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2037"
         Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(40)=   "Column(5)._HeadDivider=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Width=79"
         Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=8708"
         Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
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
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
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
         Height          =   1350
         Left            =   7365
         TabIndex        =   6
         Top             =   150
         Width           =   5610
         _Version        =   65536
         _ExtentX        =   9895
         _ExtentY        =   2381
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
      Begin TDBDate6Ctl.TDBDate txtDateFiled 
         Height          =   300
         Left            =   6090
         TabIndex        =   3
         Top             =   870
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   529
         Calendar        =   "frmLeave2.frx":83A2
         Caption         =   "frmLeave2.frx":84A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":850E
         Keys            =   "frmLeave2.frx":852C
         Spin            =   "frmLeave2.frx":858A
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
         Top             =   540
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   529
         Caption         =   "frmLeave2.frx":85B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":861E
         Key             =   "frmLeave2.frx":863C
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
      Begin TDBDate6Ctl.TDBDate txtFromDate 
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
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
      Begin TDBDate6Ctl.TDBDate txtToDate 
         Height          =   300
         Left            =   4695
         TabIndex        =   5
         Top             =   1200
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
         _ExtentY        =   529
         Calendar        =   "frmLeave2.frx":8890
         Caption         =   "frmLeave2.frx":8996
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":89FC
         Keys            =   "frmLeave2.frx":8A1A
         Spin            =   "frmLeave2.frx":8A78
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
      Begin TDBNumber6Ctl.TDBNumber txtAmount 
         Height          =   300
         Left            =   -135
         TabIndex        =   20
         Top             =   3390
         Visible         =   0   'False
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   529
         Calculator      =   "frmLeave2.frx":8AA0
         Caption         =   "frmLeave2.frx":8AC0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":8B26
         Keys            =   "frmLeave2.frx":8B44
         Spin            =   "frmLeave2.frx":8B8E
         AlignHorizontal =   1
         AlignVertical   =   2
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1990328325
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText txtLeaveType 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   870
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmLeave2.frx":8BB6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeave2.frx":8C22
         Key             =   "frmLeave2.frx":8C40
         BackColor       =   16777215
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
      Begin TrueOleDBList80.TDBCombo tdbTmp 
         Bindings        =   "frmLeave2.frx":8C84
         DataMember      =   "tdbJob"
         Height          =   300
         Left            =   135
         TabIndex        =   23
         Top             =   2145
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   529
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   0
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   0
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   0   'False
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"frmLeave2.frx":8C95
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H404040&"
         _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
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
         Left            =   -75
         TabIndex        =   22
         Top             =   915
         Width           =   1560
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2985
         TabIndex        =   19
         Top             =   1245
         Width           =   1560
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -90
         TabIndex        =   18
         Top             =   1245
         Width           =   1560
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
         TabIndex        =   16
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
         Left            =   5370
         TabIndex        =   15
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
         Left            =   300
         TabIndex        =   14
         Top             =   1650
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
         Left            =   4440
         TabIndex        =   12
         Top             =   915
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
         TabIndex        =   11
         Top             =   255
         Width           =   1560
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "modiifie"
      Height          =   375
      Left            =   13185
      TabIndex        =   9
      Top             =   3105
      Width           =   1215
   End
End
Attribute VB_Name = "frmLOBLeave2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd                 As Boolean
Dim mTxt                    As TDBText
Dim rsTmpLeaveEntry         As ADODB.Recordset
Dim mFromDate               As String
Dim mToDate                 As String

Private Sub Form_Load()

  Dim rsLeaveEntry  As ADODB.Recordset

    'Create_TmpLeaveEntry
    
    
    txtFullname.Text = frmLOBLeave.txtFullname
    
  If mAdd = False Then
    Create_TmpLeaveEntry
    With frmLOBLeave.rsLeaves
      If .RecordCount > 0 Then
        
        NetOpen rsLeaveEntry, "select x1.* from leaveapp_lines x1 " & _
                                  "where x1.leaveapp_id = '" & !leaveapp_id & "'"

        txtLvNum.Text = !leaveapp_id
        txtLeaveType.Tag = !leavetypescode
        txtLeaveType.Text = !leavetypesname
        txtDateFiled.Text = Format(!datefiled, "MM/DD/YYYY")
        txtFromDate.Text = Format(!fromdate, "MM/DD/YYYY")
        txtFromDate.Tag = Format(!fromdate, "MM/DD/YYYY")
        txtToDate.Text = Format(!todate, "MM/DD/YYYY")
        txtToDate.Tag = Format(!todate, "MM/DD/YYYY")
        txtRemarks.Text = !remarks

        If rsLeaveEntry.RecordCount > 0 Then
          rsLeaveEntry.MoveFirst
          Do While Not rsLeaveEntry.EOF
            rsTmpLeaveEntry.AddNew
'            rsTmpLeaveEntry.Fields("leaveapp_date") = rsLeaveEntry!leavetypescode
'            rsTmpLeaveEntry.Fields("leavetypesname") = rsLeaveEntry!LeaveTypesname
'            rsTmpLeaveEntry.Fields("fromdate") = rsLeaveEntry!fromdate
'            rsTmpLeaveEntry.Fields("todate") = rsLeaveEntry!todate
'            rsTmpLeaveEntry.Fields("withpay") = rsLeaveEntry!withpay
'            rsTmpLeaveEntry.Fields("firstshift") = rsLeaveEntry!firstshift
'            rsTmpLeaveEntry.Fields("secondshift") = rsLeaveEntry!secondshift
            
            rsTmpLeaveEntry.Fields("leaveapp_date") = CDate(rsLeaveEntry!leaveapp_date)
            rsTmpLeaveEntry.Fields("leaveapp_day") = Format(CDate(rsLeaveEntry!leaveapp_date), "dddd")
            rsTmpLeaveEntry.Fields("leaveapp_hours") = rsLeaveEntry!leaveapp_hours
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
    txtFromDate.Text = ""
    txtToDate.Text = ""
  End If

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorHandler

    Dim rsChk               As ADODB.Recordset
    Dim mAppDate            As Date

    Dim mProcessType        As String
    Dim mMessage            As String

    If Trim(txtLeaveType.Tag) = "" Then
        MsgBox "Please specify the type of leave.", vbExclamation + vbOKOnly
        txtLeaveType.SetFocus
        Exit Sub
    End If
    If Not IsDate(txtDateFiled.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtDateFiled.SetFocus
      Exit Sub
    End If
    If Not IsDate(txtFromDate.Tag) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtFromDate.SetFocus
      Exit Sub
    End If
    If Not IsDate(txtToDate.Tag) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtToDate.SetFocus
      Exit Sub
    End If
    If rsTmpLeaveEntry Is Nothing Then
        MsgBox "Empty leave entries.", vbExclamation + vbOKOnly
        txtFromDate.SetFocus
        Exit Sub
    End If
    If rsTmpLeaveEntry.RecordCount <= 0 Then
      MsgBox "Please verify the dates of validity for leave application.", vbExclamation + vbOKOnly
      txtFromDate.SetFocus
      Exit Sub
    End If
    If CDate(txtFromDate.Value) > CDate(txtToDate.Value) Then
        MsgBox "Invalid range of dates.", vbExclamation + vbOKOnly
        txtFromDate.SetFocus
        Exit Sub
    End If
    
'    With rsTmpLeaveEntry
'      .MoveFirst
'      Do While Not .EOF
'        If !leavetypescode <> "" Then
'
'            If Not IsDate(!fromdate) Then
'                MsgBox "Please enter a valid date."
'                'tdgLeaveEntry.SetFocus
'                Exit Sub
'            End If
'
'            If Not IsDate(!todate) Then
'                MsgBox "Please enter a valid date."
'                'tdgLeaveEntry.SetFocus
'                Exit Sub
'            End If
'
'            If CDate(Format(!fromdate, "MM/DD/YYYY")) > CDate(Format(!todate, "MM/DD/YYYY")) Then
'                MsgBox "End date is earlier than the start date.", vbExclamation + vbOKOnly
'                'tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
'                'tdgLeaveEntry.SetFocus
'                Exit Sub
'            End If
'
'            If !firstshift = 0 And !secondshift = 0 Then
'                MsgBox "You must select at least one shift.", vbExclamation + vbOKOnly
'                'tdgLeaveEntry.Col = tdgLeaveEntry.Columns("firstshift").ColIndex
'                'tdgLeaveEntry.SetFocus
'                Exit Sub
'            End If
'
'        End If
'        .MoveNext
'
'      Loop
'    End With

    If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If

    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    With frmLOBLeave
        If mAdd = True Then
          'txtLvNum.Text = LastCode("Leaves")
          ConMain.Execute "insert into leaveapp_headers (employeecode,costcentercode,divisioncode,branchcode, " & _
                                "leavetypescode,datefiled,fromdate,todate,trnxdatetime,remarks) values (" & _
                                 .mEmployeeCode & "," & .mCostCenterCode & "," & .mDivisionCode & "," & .mBranchCode & ", " & _
                                "'" & txtLeaveType.Tag & "','" & Format(txtDateFiled.Value, "YYYY-MM-DD") & "','" & Format(txtFromDate.Value, "YYYY-MM-DD") & "','" & Format(txtToDate.Value, "YYYY-MM-DD") & "', " & _
                                "now(),'" & Swap(txtRemarks.Text) & "')"
          NetOpen rsChk, "select LAST_INSERT_ID() as last_ID"
          txtLvNum.Text = rsChk!last_id
        Else
          ConMain.Execute "update leaveapp_headers set costcentercode = " & .mCostCenterCode & ", divisioncode = " & .mDivisionCode & ", branchcode = " & .mBranchCode & ", " & _
                                "leavetypescode=" & txtLeaveType.Tag & ",datefiled = '" & Format(txtDateFiled.Value, "YYYY-MM-DD") & "', updtrnxdatetime=now(), " & _
                                "fromdate='" & Format(txtFromDate.Value, "YYYY-MM-DD") & "',todate='" & Format(txtToDate.Value, "YYYY-MM-DD") & "'," & _
                                "remarks = '" & Swap(txtRemarks.Text) & "' where leaveapp_id = '" & txtLvNum.Text & "'"
          ConMain.Execute "delete from leaveapp_lines where leaveapp_id = '" & txtLvNum.Text & "'"
          'ConMain.Execute "delete from appdate where apptype = 'LV' and trnxcode = '" & txtLvNum.Text & "'"
        End If
    End With
    
    
    With rsTmpLeaveEntry
        .MoveFirst
        Do While Not .EOF
            mProcessType = "Inserting Details"
            mMessage = Format(!leaveapp_date, "MM/DD/YYYY")
            ConMain.Execute "insert into leaveapp_lines (leaveapp_id,employeecode,leavetypescode,leaveapp_date,leaveapp_hours,firstshift,secondshift,withpay) values " & _
                            "('" & txtLvNum.Text & "','" & frmLOBLeave.mEmployeeCode & "','" & txtLeaveType.Tag & "', " & _
                            "'" & Format(!leaveapp_date, "YYYY-MM-DD") & "'," & Format(!leaveapp_hours, "##0.00") & ",'" & IIf(!firstshift <> 0, 1, 0) & "','" & IIf(!secondshift <> 0, 1, 0) & "','" & IIf(!withpay <> 0, 1, 0) & "')"
            mProcessType = "Success"
'            mAppDate = Format(!fromdate, "MM/DD/YYYY")
'            Do While mAppDate <= CDate(Format(!todate, "MM/DD/YYYY"))
'                mErrorType = "Duplicate entry"
'                If !firstshift = 1 And !secondshift = 1 Then
'                    NetOpen rsChk, "select * from appdate where trnxdate = '" & Format(mAppDate, "YYYY-MM-DD") & "' and employeecode = " & frmLOBLeave.mEmployeeCode & " and (firstshift = 1 or secondshift = 1)"
'                    If rsChk.RecordCount > 0 Then
'                        If rsChk!apptype = "LV" Then
'                            MsgBox "Employee has already a leave application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
'                            ConMain.RollbackTrans
'                            'tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
'                            'tdgLeaveEntry.SetFocus
'                            Exit Sub
'                        ElseIf rsChk!apptype = "OBT" Then
'                            MsgBox "Employee has a already travel application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
'                            ConMain.RollbackTrans
'                            'tdgLeaveEntry.Col = tdgLeaveEntry.Columns("fromdate").ColIndex
'                            'tdgLeaveEntry.SetFocus
'                            Exit Sub
'                        End If
'                    End If
'                End If
'                ConMain.Execute "insert into appdate (apptype,employeecode,trnxcode,trnxdate,leavetypescode,firstshift,secondshift,withpay,processed) values ('LV'," & frmLOBLeave.mEmployeeCode & ",'" & txtLvNum.Text & "','" & Format(mAppDate, "YYYY-MM-DD") & "'," & !leavetypescode & ",'" & IIf(!firstshift <= 0, 0, 1) & "','" & IIf(!secondshift <= 0, 0, 1) & "','" & IIf(!withpay <> 0, 1, 0) & "','N')"
'                mAppDate = mAppDate + 1
'            Loop
            .MoveNext
        Loop
    End With

    

'    Exit Sub

ErrorHandler:

    If mProcessType = "Inserting Details" Then

'        With rsTmpLeaveEntry
'            NetOpen rsChk, "select * from appdate where trnxdate = '" & Format(mAppDate, "YYYY-MM-DD") & "' and employeecode = " & frmLOBLeave.mEmployeeCode & " and firstshift = " & IIf(!firstshift <= 0, 0, 1) & " and secondshift = " & IIf(!secondshift <= 0, 0, 1) & ""
'            If rsChk.RecordCount > 0 Then
'                If rsChk!apptype = "LV" Then
'                    MsgBox "Employee has already a leave application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
'                ElseIf rsChk!apptype = "OBT" Then
'                    MsgBox "Employee has a already travel application on " & Format(mAppDate, "MMMM DD,YYYY") & ".", vbExclamation + vbOKOnly, "Duplicate entry"
'                End If
'            End If
'        End With
        MsgBox "Employee has already a leave application on " & mMessage & ".", vbExclamation + vbOKOnly, "Duplicate entry"
        ConMain.RollbackTrans
    Else
        ConMain.CommitTrans
    
        frmLOBLeave.rsLeaves.Requery
        frmLOBLeave.rsLeaves.Find "leaveapp_id = '" & txtLvNum.Text & "'"
        Lock_Button "TTTT", frmLOBLeave.cmdMenu, 3
    
        Unload Me
    End If

End Sub

Private Sub Create_TmpLeaveEntry()

  Set rsTmpLeaveEntry = Nothing
  Set rsTmpLeaveEntry = New ADODB.Recordset

  With rsTmpLeaveEntry
    .Fields.Append "leaveapp_date", adDate
    .Fields.Append "leaveapp_day", adVarChar, 10
    .Fields.Append "leaveapp_hours", adDouble
    .Fields.Append "withpay", adInteger
    .Fields.Append "firstshift", adInteger
    .Fields.Append "secondshift", adInteger
    .Open
  End With

  Set tdglvDays.DataSource = rsTmpLeaveEntry

End Sub

Private Sub tdbTmp_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Trim(tdbTmp.Text) <> "" And Not IsNull(tdbTmp.SelectedItem) And tdbTmp.ApproxCount > 0 Then
      mTxt.Tag = tdbTmp.BoundText
      mTxt.Text = tdbTmp.Text
    Else
      mTxt.Tag = ""
      mTxt.Text = ""
    End If
    mTxt.Visible = True
    mTxt.SetFocus
    tdbTmp.Visible = False
  Else
    SearchList KeyAscii, tdbTmp, tdbTmp.RowSource, tdbTmp.Text
  End If
End Sub

Private Sub tdbTmp_LostFocus()
  mTxt.Visible = True
  tdbTmp.Visible = False
End Sub

Private Sub tdglvDays_AfterColUpdate(ByVal ColIndex As Integer)
    With tdglvDays
        If ColIndex = .Columns("leaveapp_hours").ColIndex Then
            If .Columns("leaveapp_hours").Value > 0 Then
                .Columns("firstshift").Value = 1
                .Columns("secondshift").Value = 1
                .Columns("withpay").Value = 1
            Else
                .Columns("firstshift").Value = 0
                .Columns("secondshift").Value = 0
                .Columns("withpay").Value = 0
            End If
        End If
    End With
End Sub

Private Sub tdglvDays_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)


  With tdglvDays

      If ColIndex = .Columns("firstshift").ColIndex Or _
         ColIndex = .Columns("secondshift").ColIndex Or _
         ColIndex = .Columns("withpay").ColIndex Then

        If .Columns("leaveapp_hours").Value <= 0 Then Cancel = True

'        If ColIndex = .Columns("firstshift").ColIndex Or ColIndex = .Columns("secondshift").ColIndex Then
'            If Format(.Columns("fromdate").Text, "MM/DD/YYYY") <> Format(.Columns("todate").Text, "MM/DD/YYYY") Then
'                Cancel = True
'            End If
'        End If
'
'        If ColIndex = .Columns("firstshift").ColIndex Then
'            If .Columns("secondshift").Value = 0 Then
'                Cancel = True
'            End If
'        End If
'
'        If ColIndex = .Columns("secondshift").ColIndex Then
'            If .Columns("firstshift").Value = 0 Then
'                Cancel = True
'            End If
'        End If
      End If

    End With

End Sub

Private Sub txtDateFiled_GotFocus()
    With txtDateFiled
        .SelStart = 0
        .SelLength = Len(txtDateFiled.Text)
    End With
End Sub

Private Sub txtDateFiled_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdLeaveType_Click()
  bind_tdb ConMain, tdbTmp, "select leavetypescode,leavetypesname from leavetypes order by leavetypesname", "leavetypesname", "leavetypescode"
  Set mTxt = txtLeaveType
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
End Sub

Private Sub txtFromDate_GotFocus()
    With txtFromDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNull(txtFromDate.Value) Then
            MsgBox "Invalid date format.", vbExclamation + vbOKOnly
            Set rsTmpLeaveEntry = Nothing
            Set tdglvDays.DataSource = Nothing
            txtFromDate.SetFocus
        Else
            Load_Leaves
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtFromDate_LostFocus()
    With txtFromDate
        If IsDate(.Value) Then
            If IsDate(.Tag) Then
                If CDate(.Tag) <> CDate(.Value) Then
                    Set rsTmpLeaveEntry = Nothing
                    Set tdglvDays.DataSource = Nothing
                End If
            End If
            .Tag = .Value
        Else
        End If
    End With
End Sub

Private Sub txtFullname_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLeaveType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
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

Private Sub txtToDate_GotFocus()
    With txtToDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNull(txtToDate.Value) Then
            MsgBox "Invalid date format.", vbExclamation + vbOKOnly
            Set rsTmpLeaveEntry = Nothing
            Set tdglvDays.DataSource = Nothing
            txtToDate.SetFocus
        Else
            Load_Leaves
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub Load_Leaves()
    If IsNull(txtFromDate.Value) Then
        MsgBox "Invalid date format.", vbExclamation + vbOKOnly
        Set rsTmpLeaveEntry = Nothing
        Set tdglvDays.DataSource = Nothing
        txtFromDate.SetFocus
        Exit Sub
    End If
    If IsNull(txtToDate.Value) Then
        txtToDate.Value = txtFromDate.Value
        txtToDate.Tag = txtFromDate.Value
    End If
    
    Dim rsLvHrParam As ADODB.Recordset
    Dim dblDay      As Double
    Dim dblStart    As Double
    Dim dblEnd      As Double
    
    dblStart = txtFromDate.Value
    dblEnd = txtToDate.Value
    Create_TmpLeaveEntry
    txtFromDate.Tag = txtFromDate.Value
    txtToDate.Tag = txtToDate.Value
    NetOpen rsLvHrParam, "select * from lvhr_parameters where lvhrparam_id=" & frmLOBLeave.mLvHrParam_ID
    For dblDay = dblStart To dblEnd Step 1
        rsTmpLeaveEntry.AddNew
        rsTmpLeaveEntry.Fields("leaveapp_date") = CDate(dblDay)
        rsTmpLeaveEntry.Fields("leaveapp_day") = Format(CDate(dblDay), "dddd")
        rsTmpLeaveEntry.Fields("leaveapp_hours") = rsLvHrParam(Weekday(CDate(dblDay))).Value
        If CDbl(rsTmpLeaveEntry!leaveapp_hours) > 0 Then
            rsTmpLeaveEntry.Fields("withpay") = 1
            rsTmpLeaveEntry.Fields("firstshift") = 1
            rsTmpLeaveEntry.Fields("secondshift") = 1
        Else
            rsTmpLeaveEntry.Fields("withpay") = 0
            rsTmpLeaveEntry.Fields("firstshift") = 0
            rsTmpLeaveEntry.Fields("secondshift") = 0
        End If
        
        rsTmpLeaveEntry.Update
    Next dblDay
End Sub
Private Sub txtToDate_LostFocus()
    With txtToDate
        If IsDate(.Value) Then
            If IsDate(.Tag) Then
                If CDate(.Tag) <> CDate(.Value) Then
                    Set rsTmpLeaveEntry = Nothing
                    Set tdglvDays.DataSource = Nothing
                End If
            End If
            .Tag = .Value
        Else
        End If
    End With
End Sub
