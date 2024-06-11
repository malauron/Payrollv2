VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADGatePasses2 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6330
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
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
   Begin TrueOleDBGrid80.TDBGrid tdgGPLine 
      Height          =   3945
      Left            =   60
      TabIndex        =   1
      Top             =   1965
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   6959
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "otlne"
      Columns(0).DataField=   "gplneno"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Date"
      Columns(1).DataField=   "datelog"
      Columns(1).NumberFormat=   "MM/DD/YYYY"
      Columns(1).ExternalEditor=   "txtDateLog"
      Columns(1).ExternalEditor.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Time"
      Columns(2).DataField=   "timelog"
      Columns(2).NumberFormat=   "hh:nn AM/PM"
      Columns(2).ExternalEditor=   "txtTime"
      Columns(2).ExternalEditor.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Log"
      Columns(3).DataField=   "logstat"
      Columns(3).NumberFormat=   "hh:nn AM/PM"
      Columns(3).DropDown=   "tddLogStatus"
      Columns(3).DropDown.vt=   8
      Columns(3).ExternalEditor=   "txtLogStat"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Remarks"
      Columns(4).DataField=   "remarks"
      Columns(4).ExternalEditor=   "txtRemarks"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Status"
      Columns(5).DataField=   "status"
      Columns(5).DropDown=   "tddStatus"
      Columns(5).DropDown.vt=   8
      Columns(5).ExternalEditor=   "txtStatus"
      Columns(5).ExternalEditor.vt=   8
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Approved By"
      Columns(6).DataField=   "approvby"
      Columns(6).DropDown=   "tddSignatory"
      Columns(6).DropDown.vt=   8
      Columns(6).ExternalEditor=   "txtSignatory"
      Columns(6).ExternalEditor.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "enteredbyuser"
      Columns(7).DataField=   "enteredbyuser"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1667"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1588"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1879"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1826"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1746"
      Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3201"
      Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=8196"
      Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=66,.parent=13,.locked=-1"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBDropDown tddSignatory 
      Height          =   1365
      Left            =   8490
      TabIndex        =   2
      Top             =   6945
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Name"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
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
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HDAFAEF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
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
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TDBTime6Ctl.TDBTime txtTime 
      Height          =   285
      Left            =   6795
      TabIndex        =   3
      Top             =   6855
      Visible         =   0   'False
      Width           =   1650
      _Version        =   65536
      _ExtentX        =   2910
      _ExtentY        =   503
      Caption         =   "frmADGatePasses2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frmADGatePasses2.frx":006C
      Spin            =   "frmADGatePasses2.frx":00BC
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
      Text            =   "02:25 PM"
      ValidateMode    =   0
      ValueVT         =   1935998983
      Value           =   0.600914351851852
   End
   Begin TDBDate6Ctl.TDBDate txtDateLog 
      Height          =   285
      Left            =   5460
      TabIndex        =   4
      Top             =   6930
      Visible         =   0   'False
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   503
      Calendar        =   "frmADGatePasses2.frx":00E4
      Caption         =   "frmADGatePasses2.frx":01FC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADGatePasses2.frx":0268
      Keys            =   "frmADGatePasses2.frx":0286
      Spin            =   "frmADGatePasses2.frx":02E4
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
      Text            =   "04/02/2008"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39540
      CenturyMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtRemarks 
      Height          =   285
      Left            =   8940
      TabIndex        =   5
      Top             =   6855
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADGatePasses2.frx":030C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADGatePasses2.frx":0378
      Key             =   "frmADGatePasses2.frx":0396
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
   Begin TrueOleDBGrid80.TDBDropDown tddStatus 
      Height          =   1365
      Left            =   3465
      TabIndex        =   6
      Top             =   7020
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Status"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
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
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HDAFAEF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
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
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText txtStatus 
      Height          =   285
      Left            =   6330
      TabIndex        =   7
      Top             =   6810
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADGatePasses2.frx":03DA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADGatePasses2.frx":0446
      Key             =   "frmADGatePasses2.frx":0464
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
   Begin TDBText6Ctl.TDBText txtSignatory 
      Height          =   285
      Left            =   3810
      TabIndex        =   8
      Top             =   6780
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADGatePasses2.frx":04A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADGatePasses2.frx":0514
      Key             =   "frmADGatePasses2.frx":0532
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
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   45
      TabIndex        =   9
      Top             =   300
      Width           =   11175
      Begin TDBText6Ctl.TDBText txtCostCenter 
         Height          =   300
         Left            =   7125
         TabIndex        =   10
         Top             =   330
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmADGatePasses2.frx":0576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADGatePasses2.frx":05E2
         Key             =   "frmADGatePasses2.frx":0600
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
      Begin TDBText6Ctl.TDBText txtDivision 
         Height          =   300
         Left            =   7125
         TabIndex        =   11
         Top             =   660
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmADGatePasses2.frx":0644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADGatePasses2.frx":06B0
         Key             =   "frmADGatePasses2.frx":06CE
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
      Begin TDBText6Ctl.TDBText txtBranch 
         Height          =   300
         Left            =   7125
         TabIndex        =   12
         Top             =   990
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmADGatePasses2.frx":0712
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADGatePasses2.frx":077E
         Key             =   "frmADGatePasses2.frx":079C
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
      Begin TrueOleDBList80.TDBCombo tdbEmpNo 
         Height          =   315
         Left            =   1755
         TabIndex        =   13
         Top             =   660
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
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
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _PropDict       =   $"frmADGatePasses2.frx":07E0
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
      Begin TrueOleDBList80.TDBCombo tdbEmpName 
         Height          =   345
         Left            =   1755
         TabIndex        =   14
         Tag             =   "Municipal"
         Top             =   1005
         Width           =   3000
         _ExtentX        =   5292
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
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "costcentercode"
         Columns(2).DataField=   "costcentercode"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "divisioncode"
         Columns(3).DataField=   "divisioncode"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "branchcode"
         Columns(4).DataField=   "branchcode"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "costcenter"
         Columns(5).DataField=   "costcenter"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "division"
         Columns(6).DataField=   "division"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "branch"
         Columns(7).DataField=   "branch"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
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
         _PropDict       =   $"frmADGatePasses2.frx":088A
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin TDBText6Ctl.TDBText txtGPCode 
         Height          =   300
         Left            =   1755
         TabIndex        =   15
         Top             =   330
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   529
         Caption         =   "frmADGatePasses2.frx":0934
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADGatePasses2.frx":09A0
         Key             =   "frmADGatePasses2.frx":09BE
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Branch"
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
         Height          =   240
         Left            =   5595
         TabIndex        =   21
         Top             =   1065
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Division"
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
         Height          =   240
         Left            =   5595
         TabIndex        =   20
         Top             =   705
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cost Center"
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
         Height          =   240
         Left            =   5610
         TabIndex        =   19
         Top             =   345
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H0030A0B8&
         Height          =   240
         Left            =   255
         TabIndex        =   18
         Top             =   1095
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee ID"
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
         Height          =   240
         Left            =   255
         TabIndex        =   17
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Number"
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
         Height          =   240
         Left            =   255
         TabIndex        =   16
         Top             =   375
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0030A0B8&
         X1              =   5415
         X2              =   5415
         Y1              =   240
         Y2              =   1500
      End
   End
   Begin VB.Frame fraButton 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   30
      TabIndex        =   22
      Top             =   5895
      Width           =   3645
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   330
         Left            =   30
         TabIndex        =   23
         Top             =   60
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         Caption         =   "&Save"
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
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   330
         Left            =   1815
         TabIndex        =   24
         Top             =   60
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
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
         cFore           =   3186872
         cFHover         =   3186872
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
   End
   Begin TrueOleDBGrid80.TDBDropDown tddLogStatus 
      Height          =   1365
      Left            =   270
      TabIndex        =   25
      Top             =   6780
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2408
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Code"
      Columns(0).DataField=   "code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Log Status"
      Columns(1).DataField=   "description"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
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
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(14)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HDAFAEF&"
      _StyleDefs(16)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
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
      _StyleDefs(31)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(32)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(33)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(35)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(39)  =   "Named:id=33:Normal"
      _StyleDefs(40)  =   ":id=33,.parent=0"
      _StyleDefs(41)  =   "Named:id=34:Heading"
      _StyleDefs(42)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(43)  =   ":id=34,.wraptext=-1"
      _StyleDefs(44)  =   "Named:id=35:Footing"
      _StyleDefs(45)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   "Named:id=36:Selected"
      _StyleDefs(47)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(48)  =   "Named:id=37:Caption"
      _StyleDefs(49)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(50)  =   "Named:id=38:HighlightRow"
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText txtLogStat 
      Height          =   285
      Left            =   5145
      TabIndex        =   26
      Top             =   6270
      Visible         =   0   'False
      Width           =   3000
      _Version        =   65536
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "frmADGatePasses2.frx":0A02
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmADGatePasses2.frx":0A6E
      Key             =   "frmADGatePasses2.frx":0A8C
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
End
Attribute VB_Name = "frmADGatePasses2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd         As Boolean
Dim rsTempGPEntry   As ADODB.Recordset

Private Sub Form_Load()

  Dim rsTmp         As ADODB.Recordset

  Dim I             As Integer

  CreateTmpDB rsTmp

  With rsTmp
    For I = 1 To 3
        .AddNew
        Select Case I
            Case 1: .Fields("code") = "Approved"
                    .Fields("description") = "Approved"
            Case 2: .Fields("code") = "Cancelled"
                    .Fields("description") = "Cancelled"
            Case 3: .Fields("code") = "logstatimelogg"
                    .Fields("description") = "logstatimelogg"
        End Select
        .Update
    Next
  End With

  With tddStatus
    .DataSource = rsTmp
    .ListField = "description"
  End With

  Set rsTmp = Nothing
  
  CreateTmpDB rsTmp

  With rsTmp
    For I = 1 To 2
        .AddNew
        Select Case I
            Case 1: .Fields("code") = "In"
                    .Fields("description") = "In"
            Case 2: .Fields("code") = "Out"
                    .Fields("description") = "Out"
        End Select
        .Update
    Next
  End With

  With tddLogStatus
    .DataSource = rsTmp
    .ListField = "description"
  End With


  bind_tdb ConMain, tdbEmpNo, "select employeecode,dummycode from employee " & _
                            "where payfreqcode = '" & frmADGatePasses.tdbPayrollPeriod.Columns("payfreqcode").Text & "' " & _
                            "order by dummycode", "dummycode", "employeecode"

  bind_tdb ConMain, tdbEmpName, "select x1.employeecode,concat(x1.lastname,', ',x1.firstname,' ',x1.middlename) fullname, " & _
                          "x1.costcentercode, x1.divisioncode, x1.branchcode,x2.costcenter,x3.division,x4.branch from employee x1 " & _
                          "left outer join costcenter x2 on x1.costcentercode = x2.costcentercode " & _
                          "left outer join division x3 on x1.divisioncode = x3.divisioncode " & _
                          "left outer join branch x4 on x1.branchcode = x4.branchcode " & _
                          "where payfreqcode = '" & frmADGatePasses.tdbPayrollPeriod.Columns("payfreqcode").Text & "' " & _
                          "order by x1.lastname,x1.firstname,x1.middlename", "fullname", "employeecode"

  Bind_tdd ConMain, tddSignatory, "select fullname code,fullname description from signatory order by fullname", "fullname"

  If mAdd = False Then
    With frmADGatePasses.rsGatePass
      If .RecordCount > 0 Then
        txtGPCode.Text = !gpcode
        tdbEmpNo.BoundText = !employeecode
        tdbEmpName.BoundText = !employeecode
        tdbEmpNo.Enabled = False
        tdbEmpName.Enabled = False
        txtCostCenter.Text = !CostCenter & ""
        txtDivision.Text = !Division & ""
        txtBranch.Text = !branch & ""
        Load_OT !employeecode
      End If
    End With
  Else
'    txtDateFiled.Text = ""
  End If

End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim rsChk       As ADODB.Recordset
    
    Dim mgpcode     As String
    Dim mPerCode    As String
    Dim mDate       As String
    Dim mEntrdBU    As String
    Dim mgplneno    As String

    If Trim(tdbEmpName.Text) = "" Or IsNull(tdbEmpName.SelectedItem) Or tdbEmpName.ApproxCount = 0 Then
      MsgBox "Please select an employee.", vbExclamation + vbOKOnly
      tdbEmpName.SetFocus
      Exit Sub
    End If
    
    If rsTempGPEntry Is Nothing Then
      MsgBox "You must provide at least one (1) overtime entry.", vbExclamation + vbOKOnly
      tdgGPLine.SetFocus
      Exit Sub
    End If

    If rsTempGPEntry.RecordCount <= 0 Then
      MsgBox "You must provide at least one (1) overtime entry.", vbExclamation + vbOKOnly
      tdgGPLine.SetFocus
      Exit Sub
    End If

    With rsTempGPEntry
      .MoveFirst
      Do While Not .EOF
        If Not IsDate(!datelog) Then
          MsgBox "Please enter a valid date."
          tdgGPLine.SetFocus
          Exit Sub
        End If
        .MoveNext
      Loop
    End With

    If MsgBox("Confirm saving data.", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans

    If mAdd = True Then
    
        If isEmpExists(tdbEmpNo.Text) = True Then
            MsgBox "Employee has already an existimelogg overime entry.", vbExclamation + vbOKOnly
            ConMain.RollbackTrans
            Exit Sub
        End If
    
        txtGPCode.Text = LastCode("gphdr")
        
        ConMain.Execute "insert into gphdr(gpcode,employeecode,percode,costcentercode,divisioncode,branchcode,payyear,paymonth,tdatetime) values " & _
                              "('" & txtGPCode.Text & "','" & tdbEmpNo.Text & "','" & frmADGatePasses.tdbPayrollPeriod.Columns("percode").Text & "', " & _
                              "'" & tdbEmpName.Columns("costcentercode").Text & "', " & _
                              "'" & tdbEmpName.Columns("divisioncode").Text & "','" & tdbEmpName.Columns("branchcode").Text & "', " & _
                              "'" & frmADGatePasses.tdbPayrollPeriod.Columns("payyear").Text & "', " & _
                              "'" & frmADGatePasses.tdbPayrollPeriod.Columns("paymonth").Text & "',now())"
    End If
    
    ConMain.Execute "delete from gplne where employeecode = '" & tdbEmpNo.Text & "' and fnlz = 'N' and percode = '" & frmADGatePasses.tdbPayrollPeriod.Columns("percode").Text & "'"

    With rsTempGPEntry
        .MoveFirst
        Do While Not .EOF
        
        If !Status <> "logstatimelogg" Then
            mgpcode = txtGPCode.Text
            mPerCode = frmADGatePasses.tdbPayrollPeriod.Columns("percode").Text
        Else
            mgpcode = ""
            mPerCode = ""
        End If
        
        If Trim(!tdatetime) = "" Then
            mDate = Format(Now, "YYYY-MM-DD hh:nn:ss")
        Else
            mDate = Format(!tdatetime, "YYYY-MM-DD hh:nn:ss")
        End If
        
        If Trim(!enteredbyuser) = "" Then
            mEntrdBU = "N"
        Else
            mEntrdBU = !enteredbyuser
        End If
        
        If Trim(!gplneno) = "" Then
            mgplneno = LastCode("gplne")
        Else
            mgplneno = !gplneno
        End If
        
        ConMain.Execute "insert into gplne(gplneno,employeecode,gpcode,percode,datelog,approvby, " & _
                              "remarks,complog,timelog,logstat, " & _
                              "status,tdatetime,enteredbyuser,fnlz) values " & _
                              "('" & mgplneno & "','" & tdbEmpNo.Text & "','" & mgpcode & "','" & mPerCode & "','" & Format(!datelog, "YYYY-MM-DD") & "','" & !approvby & "', " & _
                              "'" & !remarks & "','" & Format(!datelog & " " & !timelog, "YYYY-MM-DD HH:NN:SS") & "','" & Format(!timelog, "hh:nn") & "','" & Format(!logstat, "hh:nn") & "', " & _
                              "'" & !Status & "','" & mDate & "','" & mEntrdBU & "','N')"
          
        .MoveNext
        Loop
    End With
    
    ConMain.CommitTrans

    frmADGatePasses.rsGatePass.Requery
    frmADGatePasses.rsGatePass.Find "gpcode = '" & txtGPCode.Text & "'"
    Lock_Button "TTFFTT", frmADGatePasses.cmdMenu, 5
    
    Unload Me

End Sub

Private Sub Form_Resize()
    TitleBar.Move 0, 0, Me.ScaleWidth
End Sub

Private Sub tdbEmpName_GotFocus()
    If Trim(tdbEmpName.Text) <> "" And Not IsNull(tdbEmpName.SelectedItem) And tdbEmpName.ApproxCount > 0 Then
        tdbEmpName.Tag = tdbEmpName.BoundText
    Else
        tdbEmpName.Tag = ""
    End If
End Sub

Private Sub tdbEmpName_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEmpName, tdbEmpName.RowSource, tdbEmpName.Text
  End If
End Sub

Private Sub tdbEmpName_LostFocus()
    If tdbEmpName.ApproxCount > 0 And Trim(tdbEmpName.Text) <> "" And Not IsNull(tdbEmpName.SelectedItem) Then
        tdbEmpNo.BoundText = tdbEmpName.BoundText
        txtCostCenter.Text = tdbEmpName.Columns("costcenter").Text
        txtDivision.Text = tdbEmpName.Columns("division").Text
        txtBranch.Text = tdbEmpName.Columns("branch").Text
        If isEmpExists(tdbEmpName.BoundText) = False Then
            If tdbEmpName.Tag <> tdbEmpName.BoundText Then
                Load_OT tdbEmpName.BoundText
            End If
        Else
            MsgBox "This employee already has a pull out entry.", vbExclamation + vbOKOnly
        End If
    Else
        tdbEmpNo.BoundText = ""
        Set tdgGPLine.DataSource = Nothing
    End If
End Sub

Private Sub tdbEmpno_GotFocus()
    If Trim(tdbEmpNo.Text) <> "" And Not IsNull(tdbEmpNo.SelectedItem) And tdbEmpNo.ApproxCount > 0 Then
        tdbEmpNo.Tag = tdbEmpNo.BoundText
    Else
        tdbEmpNo.Tag = ""
    End If
End Sub

Private Sub tdbEmpNo_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEmpNo, tdbEmpNo.RowSource, tdbEmpNo.Text
    tdbEmpNo_LostFocus
  End If
End Sub

Private Sub tdbEmpNo_LostFocus()
    If tdbEmpNo.ApproxCount > 0 And Trim(tdbEmpNo.Text) <> "" And Not IsNull(tdbEmpNo.SelectedItem) Then
        tdbEmpName.BoundText = tdbEmpNo.BoundText
        txtCostCenter.Text = tdbEmpName.Columns("costcenter").Text
        txtDivision.Text = tdbEmpName.Columns("division").Text
        txtBranch.Text = tdbEmpName.Columns("branch").Text
        If isEmpExists(tdbEmpNo.BoundText) = False Then
            If tdbEmpNo.Tag <> tdbEmpNo.BoundText Then
                Load_OT tdbEmpNo.BoundText
            End If
        Else
            MsgBox "This employee already has a pull out entry.", vbExclamation + vbOKOnly
        End If
    Else
        tdbEmpName.BoundText = ""
        Set tdgGPLine.DataSource = Nothing
    End If
End Sub

Private Sub tddLogStatus_RowChange()
    With tddLogStatus
        txtLogStat.Text = .Columns("description").Text
    End With
End Sub

Private Sub tddLogStatus_DropDownOpen()
    With tddLogStatus
        .Width = tdgGPLine.Columns("logstat").Width
    End With
End Sub

Private Sub tdggpline_AfterColUpdate(ByVal ColIndex As Integer)
'    With tdggpline
'        If ColIndex = .Columns("timelog").ColIndex Or ColIndex = .Columns("logstat").ColIndex Then
'            GetTtlHrs
'        End If
'    End With
End Sub

Private Sub tdggpline_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

'  If ColIndex = 2 Or ColIndex = 3 Or ColIndex = 4 Then
'    If tdggpline.Columns("leavetypescode").Text = "" Then Cancel = True
'    If ColIndex = 3 Then
'      If tdggpline.Columns("secondshift").Value = 0 Then Cancel = True
'    End If
'    If ColIndex = 4 Then
'      If tdggpline.Columns("firstshift").Value = 0 Then Cancel = True
'    End If
'  End If

End Sub

Private Sub tdggpline_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  
    With tdgGPLine
        If ColIndex = .Columns("datelog").ColIndex Then
            If IsDate(.Columns("datelog").Text) Then
                If Trim(.Columns("status").Text) = "" Then
                    .Columns("status").Text = "Approved"
                End If
            End If
        End If
    End With
    
End Sub

Private Sub tdggpline_BeforeRowColChange(Cancel As Integer)

'  If tdggpline.Columns("leavetypescode").Text <> "" Then
'
'    If tdggpline.Columns("leavetypes").Text = "" Then
'      Cancel = True
'      tdggpline.Col = tdggpline.Columns("leavetypes").ColIndex
'      SendKeys " "
'      Exit Sub
'    End If

'    If tdggpline.Columns("lvdate").Text = "" Then
'      Cancel = True
'      tdggpline.Col = tdggpline.Columns("lvdate").ColIndex
'      Exit Sub
'    End If

'    If tdggpline.Columns("firstshift").Value = 0 And tdggpline.Columns("secondshift").Value = 0 Then
'      tdggpline.Columns("firstshift").Value = 1
'      tdggpline.Columns("secondshift").Value = 1
'    End If
'
'  End If

End Sub

Private Sub tdggpline_Error(ByVal DataError As Integer, Response As Integer)
    Response = False
End Sub

Private Sub tdggpline_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If txtDateLog.Visible = False Then
            With tdgGPLine
                If rsTempGPEntry.RecordCount > 0 Then
                    .Update
                    If Not .EOF And Not .BOF Then
                      If Trim(.Columns("gplneno").Text) = "" Then
                        .Delete
                      End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub tddSignatory_RowChange()
  With tddSignatory
    txtSignatory.Text = .Columns("description").Text
  End With
End Sub

Private Sub tddSignatory_DropDownOpen()
    With tddSignatory
        .Width = tdgGPLine.Width - tdgGPLine.Columns("approvby").Left
    End With
    Bind_tdd ConMain, tddSignatory, "select fullname code,fullname description from signatory order by fullname", "description"
End Sub

Private Sub txtLogStat_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tdgGPLine.SetFocus
  Else
    SearchRecord KeyAscii, txtLogStat, tddLogStatus.DataSource, txtLogStat.Text, "description"
    tddSignatory_RowChange
  End If

End Sub

Private Sub txtLogStat_LostFocus()
    tdgGPLine.SetFocus
End Sub

Private Sub txtRemarks_LostFocus()
    tdgGPLine.SetFocus
End Sub

Private Sub txtSignatory_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    tdgGPLine.SetFocus
  Else
    SearchRecord KeyAscii, txtSignatory, tddSignatory.DataSource, txtSignatory.Text, "description"
    tddSignatory_RowChange
  End If
End Sub

Private Sub txtSignatory_LostFocus()
    tdgGPLine.SetFocus
End Sub

Private Sub tddStatus_RowChange()
    With tddStatus
        txtStatus.Text = .Columns("description").Text
    End With
End Sub

Private Sub tddStatus_DropDownOpen()
    With tddStatus
        .Width = tdgGPLine.Columns("status").Width
    End With
End Sub

Private Sub txtStatus_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tdgGPLine.SetFocus
    Else
        SearchRecord KeyAscii, txtStatus, tddStatus.DataSource, txtStatus.Text, "description"
        tddStatus_RowChange
    End If
End Sub

Private Sub txtStatus_LostFocus()
    tdgGPLine.SetFocus
End Sub

Private Sub Create_TmpOTEntry()

    If Not rsTempGPEntry Is Nothing Then
        If rsTempGPEntry.RecordCount > 0 Then
            rsTempGPEntry.MoveFirst
            Do While Not rsTempGPEntry.EOF
                rsTempGPEntry.Delete
                rsTempGPEntry.Update
                If rsTempGPEntry.RecordCount > 0 Then
                    rsTempGPEntry.MoveFirst
                Else
                    Exit Do
                End If
            Loop
        End If
    End If
    
    Set rsTempGPEntry = Nothing
    Set rsTempGPEntry = New ADODB.Recordset
    
    With rsTempGPEntry
        .Fields.Append "gplneno", adVarChar, 7
        .Fields.Append "gpcode", adVarChar, 7
        .Fields.Append "percode", adVarChar, 7
        .Fields.Append "datelog", adDate
        .Fields.Append "approvby", adVarChar, 70
        .Fields.Append "remarks", adVarChar, 100
        .Fields.Append "timelog", adVarChar, 11
        .Fields.Append "logstat", adVarChar, 11
        .Fields.Append "status", adVarChar, 20
        .Fields.Append "enteredbyuser", adVarChar, 1
        .Fields.Append "tdatetime", adVarChar, 20
        .Open
    End With
    Set tdgGPLine.DataSource = rsTempGPEntry

End Sub

Private Sub Load_OT(mEmpNo As String)

    Dim rsOTEntry     As ADODB.Recordset

    Create_TmpOTEntry
    
    NetOpen rsOTEntry, "select * from gplne where employeecode = '" & mEmpNo & "' and fnlz = 'N' and percode = '" & frmADGatePasses.tdbPayrollPeriod.Columns("percode").Text & "'"

    With rsOTEntry
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            rsTempGPEntry.AddNew
            rsTempGPEntry.Fields("gplneno") = !gplneno
            rsTempGPEntry.Fields("gpcode") = !gpcode
            rsTempGPEntry.Fields("percode") = !percode
            rsTempGPEntry.Fields("datelog") = !datelog
            rsTempGPEntry.Fields("approvby") = !approvby
            rsTempGPEntry.Fields("remarks") = !remarks
            rsTempGPEntry.Fields("timelog") = !timelog
            rsTempGPEntry.Fields("logstat") = !logstat
            rsTempGPEntry.Fields("status") = !Status
            rsTempGPEntry.Fields("enteredbyuser") = !enteredbyuser
            rsTempGPEntry.Fields("tdatetime") = Format(!tdatetime, "YYYY-MM-DD hh:nn:ss")
            rsTempGPEntry.Update
            .MoveNext
          Loop
        End If
    End With
  
End Sub

Private Function isEmpExists(mEmpNo As String) As Boolean

    Dim rsTmpOT     As ADODB.Recordset
    Set rsTmpOT = New ADODB.Recordset
    With rsTmpOT
        Set rsTmpOT.DataSource = frmADGatePasses.rsGatePass.Clone
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "employeecode = '" & mEmpNo & "'"
            If .EOF Then
                isEmpExists = False
            Else
                isEmpExists = True
            End If
        End If
    End With
    Set rsTmpOT = Nothing

End Function

Private Sub txtTime_LostFocus()
    tdgGPLine.SetFocus
End Sub


Private Sub txtdatelog_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tdgGPLine.SetFocus
    End If
End Sub

Private Sub txtdatelog_LostFocus()
    tdgGPLine.SetFocus
End Sub

