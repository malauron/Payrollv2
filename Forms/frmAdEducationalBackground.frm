VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmAdEducationalBackground 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Educational Background"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   3540
      Width           =   12825
   End
   Begin TrueOleDBGrid80.TDBGrid tdgEducationalBackground 
      Height          =   3525
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   6218
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "School Attended"
      Columns(0).DataField=   "schoolattended"
      Columns(0).ExternalEditor=   "txt"
      Columns(0).ExternalEditor.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Address"
      Columns(1).DataField=   "schooladdress"
      Columns(1).ExternalEditor=   "txt"
      Columns(1).ExternalEditor.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Level"
      Columns(2).DataField=   "schoollevel"
      Columns(2).DropDown=   "tddSchoolLevel"
      Columns(2).DropDown.vt=   8
      Columns(2).ExternalEditor=   "txtSchoolLevel"
      Columns(2).ExternalEditor.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Course"
      Columns(3).DataField=   "coursedescription"
      Columns(3).ExternalEditor=   "txt"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "From"
      Columns(4).DataField=   "fromyear"
      Columns(4).NumberFormat=   "MM-DD-YY"
      Columns(4).ExternalEditor=   "txtDate"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "To"
      Columns(5).DataField=   "toyear"
      Columns(5).NumberFormat=   "MM-DD-YY"
      Columns(5).ExternalEditor=   "txtDate"
      Columns(5).ExternalEditor.vt=   8
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=6615"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6535"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4630"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4551"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2487"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2408"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=4736"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4657"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1720"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1640"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1508"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerStyle=0"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1455"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(5)._HeadDivider=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=79"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
   Begin TDBText6Ctl.TDBText txtSchoolLevel 
      Height          =   300
      Left            =   8940
      TabIndex        =   1
      Top             =   3285
      Visible         =   0   'False
      Width           =   3060
      _Version        =   65536
      _ExtentX        =   5397
      _ExtentY        =   529
      Caption         =   "frmAdEducationalBackground.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdEducationalBackground.frx":006C
      Key             =   "frmAdEducationalBackground.frx":008A
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
   Begin TrueOleDBGrid80.TDBDropDown tddSchoolLevel 
      Height          =   1365
      Left            =   8940
      TabIndex        =   2
      Top             =   3600
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
      Splits.Count    =   1
      AllowRowSizing  =   -1  'True
      Appearance      =   2
      BorderStyle     =   1
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
      _StyleDefs(51)  =   ":id=38,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=39:EvenRow"
      _StyleDefs(53)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(54)  =   "Named:id=40:OddRow"
      _StyleDefs(55)  =   ":id=40,.parent=33"
      _StyleDefs(56)  =   "Named:id=41:RecordSelector"
      _StyleDefs(57)  =   ":id=41,.parent=34"
      _StyleDefs(58)  =   "Named:id=42:FilterBar"
      _StyleDefs(59)  =   ":id=42,.parent=33"
   End
   Begin TDBDate6Ctl.TDBDate txtDate 
      Height          =   300
      Left            =   12015
      TabIndex        =   3
      Top             =   3285
      Visible         =   0   'False
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   529
      Calendar        =   "frmAdEducationalBackground.frx":00CE
      Caption         =   "frmAdEducationalBackground.frx":01D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdEducationalBackground.frx":023A
      Keys            =   "frmAdEducationalBackground.frx":0258
      Spin            =   "frmAdEducationalBackground.frx":02B6
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
   Begin TDBText6Ctl.TDBText txt 
      Height          =   300
      Left            =   5865
      TabIndex        =   4
      Top             =   3285
      Visible         =   0   'False
      Width           =   3060
      _Version        =   65536
      _ExtentX        =   5397
      _ExtentY        =   529
      Caption         =   "frmAdEducationalBackground.frx":02DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAdEducationalBackground.frx":034A
      Key             =   "frmAdEducationalBackground.frx":0368
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
End
Attribute VB_Name = "frmAdEducationalBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTmp         As ADODB.Recordset

Private Sub Form_Load()
  
  Dim i             As Integer
  
  CreateTmpDB rsTmp
    
  For i = 1 To 4
      rsTmp.AddNew
      Select Case i
          Case 1:
            rsTmp.Fields("code") = "Elementary"
            rsTmp.Fields("description") = "Elementary"
          Case 2:
            rsTmp.Fields("code") = "High School"
            rsTmp.Fields("description") = "High School"
          Case 3:
            rsTmp.Fields("code") = "College"
            rsTmp.Fields("description") = "College"
          Case 4:
            rsTmp.Fields("code") = "Post Grad"
            rsTmp.Fields("description") = "Post Grad"
                          
      End Select
      rsTmp.Update
  Next
  Set tddSchoolLevel.DataSource = rsTmp
  tddSchoolLevel.ListField = "description"

End Sub

Private Sub tdgEducationalBackground_KeyDown(KeyCode As Integer, Shift As Integer)
  
  With tdgEducationalBackground
            If txt.Visible = False And txtDate.Visible = False And txtSchoolLevel.Visible = False Then
                If .ApproxCount > 0 Then
                    If Not .EOF Then
                        If KeyCode = 46 Then
                            If MsgBox("Do you want to delete this line.", vbQuestion + vbYesNo) = vbYes Then
                              .Delete
                              .Refresh
                            End If
                            .SetFocus
                        End If
                    End If
                End If
                
            End If
    End With
    
End Sub

Private Sub txtSchoolLevel_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tdgEducationalBackground.SetFocus
    Else
        SearchRecord KeyAscii, txtSchoolLevel, tddSchoolLevel.DataSource, txtSchoolLevel.Text, tddSchoolLevel.ListField
        'tddSchoolLevel_RowChange
    End If
End Sub

Private Sub txtSchoolLevel_LostFocus()
    tdgEducationalBackground.SetFocus
End Sub

