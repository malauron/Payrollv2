VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDWithhold 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab tabWT 
      Height          =   6495
      Left            =   105
      TabIndex        =   35
      Top             =   1125
      Width           =   9675
      _cx             =   17066
      _cy             =   11456
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16185592
      ForeColor       =   -2147483630
      FrontTabColor   =   16185592
      BackTabColor    =   16185592
      TabOutlineColor =   12632256
      FrontTabForeColor=   0
      Caption         =   "Maintain W/Tax Table|View W/Tax Table"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   4
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   6180
         Left            =   10290
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   300
         Width           =   9645
         _cx             =   17013
         _cy             =   10901
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16185592
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin TrueOleDBGrid80.TDBGrid gridWT 
            Height          =   5010
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   8325
            _ExtentX        =   14684
            _ExtentY        =   8837
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Code"
            Columns(0).DataField=   "WTCode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Description"
            Columns(1).DataField=   "Description"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Exemption"
            Columns(2).DataField=   "Exemption"
            Columns(2).NumberFormat=   "#,#0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Bracket 1"
            Columns(3).DataField=   "B1"
            Columns(3).NumberFormat=   "#,#0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Bracket 2"
            Columns(4).DataField=   "B2"
            Columns(4).NumberFormat=   "#,#0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Bracket 3"
            Columns(5).DataField=   "B3"
            Columns(5).NumberFormat=   "#,#0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Bracket 4"
            Columns(6).DataField=   "B4"
            Columns(6).NumberFormat=   "#,#0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Bracket 5"
            Columns(7).DataField=   "B5"
            Columns(7).NumberFormat=   "#,#0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Bracket 6"
            Columns(8).DataField=   "B6"
            Columns(8).NumberFormat=   "#,#0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Bracket 7"
            Columns(9).DataField=   "B7"
            Columns(9).NumberFormat=   "#,#0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Bracket 8"
            Columns(10).DataField=   "B8"
            Columns(10).NumberFormat=   "#,#0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Bracket 9"
            Columns(11).DataField=   "B9"
            Columns(11).NumberFormat=   "#,#0.00"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Bracket 10"
            Columns(12).DataField=   "B10"
            Columns(12).NumberFormat=   "#,#0.00"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1746"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(44)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(46)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(52)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(55)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(56)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(57)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(58)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(61)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=2"
            Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(64)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(67)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(68)=   "Column(11)._ColStyle=2"
            Splits(0)._ColumnProps(69)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(70)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(71)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(73)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(74)=   "Column(12)._ColStyle=2"
            Splits(0)._ColumnProps(75)=   "Column(12).Order=13"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
            _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1"
            _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.alignment=1"
            _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(86)  =   "Named:id=33:Normal"
            _StyleDefs(87)  =   ":id=33,.parent=0"
            _StyleDefs(88)  =   "Named:id=34:Heading"
            _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(90)  =   ":id=34,.wraptext=-1"
            _StyleDefs(91)  =   "Named:id=35:Footing"
            _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(93)  =   "Named:id=36:Selected"
            _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(95)  =   "Named:id=37:Caption"
            _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(97)  =   "Named:id=38:HighlightRow"
            _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(99)  =   "Named:id=39:EvenRow"
            _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(101) =   "Named:id=40:OddRow"
            _StyleDefs(102) =   ":id=40,.parent=33"
            _StyleDefs(103) =   "Named:id=41:RecordSelector"
            _StyleDefs(104) =   ":id=41,.parent=34"
            _StyleDefs(105) =   "Named:id=42:FilterBar"
            _StyleDefs(106) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBList80.TDBCombo dcboFilter 
            Height          =   300
            Left            =   930
            TabIndex        =   0
            Top             =   195
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   529
            _LayoutType     =   0
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
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _PropDict       =   $"frmMDWithhold.frx":0000
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Filter"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   660
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   6180
         Left            =   10590
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   300
         Width           =   9645
         _cx             =   17013
         _cy             =   10901
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame Frame3 
            Height          =   1350
            Left            =   180
            TabIndex        =   39
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   40
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDWithhold.frx":00AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0116
               Key             =   "frmMDWithhold.frx":0134
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
            Begin TDBText6Ctl.TDBText TDBText9 
               Height          =   300
               Left            =   1800
               TabIndex        =   41
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDWithhold.frx":0178
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":01E4
               Key             =   "frmMDWithhold.frx":0202
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
            Begin TDBText6Ctl.TDBText TDBText10 
               Height          =   300
               Left            =   1800
               TabIndex        =   42
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDWithhold.frx":0246
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":02B2
               Key             =   "frmMDWithhold.frx":02D0
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
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mode of Payment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   30
               TabIndex        =   45
               Top             =   615
               Width           =   1635
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   345
               TabIndex        =   44
               Top             =   300
               Width           =   1335
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   660
               TabIndex        =   43
               Top             =   945
               Width           =   990
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   46
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "frmMDWithhold.frx":0314
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDWithhold.frx":0380
            Key             =   "frmMDWithhold.frx":039E
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   3600
            Left            =   180
            TabIndex        =   47
            Top             =   2115
            Width           =   7275
            _cx             =   12832
            _cy             =   6350
            Appearance      =   3
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   13431287
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmMDWithhold.frx":03E2
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   915
            TabIndex        =   48
            Top             =   240
            Width           =   915
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerMAType 
         Height          =   6180
         Left            =   15
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   300
         Width           =   9645
         _cx             =   17013
         _cy             =   10901
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16185592
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame frmeWT 
            BackColor       =   &H00F6F8F8&
            Enabled         =   0   'False
            Height          =   5475
            Left            =   135
            TabIndex        =   50
            Top             =   60
            Width           =   8250
            Begin TDBNumber6Ctl.TDBNumber txtB1 
               Height          =   300
               Left            =   1005
               TabIndex        =   5
               Top             =   1470
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":044A
               Caption         =   "frmMDWithhold.frx":046A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":04D6
               Keys            =   "frmMDWithhold.frx":04F4
               Spin            =   "frmMDWithhold.frx":053E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBText6Ctl.TDBText txtWTCode 
               Height          =   300
               Left            =   1410
               TabIndex        =   2
               Top             =   270
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDWithhold.frx":0566
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":05D2
               Key             =   "frmMDWithhold.frx":05F0
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
            Begin TDBText6Ctl.TDBText txtWTDescription 
               Height          =   300
               Left            =   3645
               TabIndex        =   4
               Top             =   615
               Width           =   4440
               _Version        =   65536
               _ExtentX        =   7832
               _ExtentY        =   529
               Caption         =   "frmMDWithhold.frx":0634
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":06A0
               Key             =   "frmMDWithhold.frx":06BE
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
               ErrorBeep       =   1
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
            Begin TDBNumber6Ctl.TDBNumber txtF1 
               Height          =   300
               Left            =   3135
               TabIndex        =   6
               Top             =   1470
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0702
               Caption         =   "frmMDWithhold.frx":0722
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":078E
               Keys            =   "frmMDWithhold.frx":07AC
               Spin            =   "frmMDWithhold.frx":07F6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA1 
               Height          =   300
               Left            =   5265
               TabIndex        =   7
               Top             =   1470
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":081E
               Caption         =   "frmMDWithhold.frx":083E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":08AA
               Keys            =   "frmMDWithhold.frx":08C8
               Spin            =   "frmMDWithhold.frx":0912
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB2 
               Height          =   300
               Left            =   1005
               TabIndex        =   8
               Top             =   1860
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":093A
               Caption         =   "frmMDWithhold.frx":095A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":09C6
               Keys            =   "frmMDWithhold.frx":09E4
               Spin            =   "frmMDWithhold.frx":0A2E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF2 
               Height          =   300
               Left            =   3135
               TabIndex        =   9
               Top             =   1860
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0A56
               Caption         =   "frmMDWithhold.frx":0A76
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0AE2
               Keys            =   "frmMDWithhold.frx":0B00
               Spin            =   "frmMDWithhold.frx":0B4A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA2 
               Height          =   300
               Left            =   5265
               TabIndex        =   10
               Top             =   1860
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0B72
               Caption         =   "frmMDWithhold.frx":0B92
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0BFE
               Keys            =   "frmMDWithhold.frx":0C1C
               Spin            =   "frmMDWithhold.frx":0C66
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB3 
               Height          =   300
               Left            =   1005
               TabIndex        =   11
               Top             =   2250
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0C8E
               Caption         =   "frmMDWithhold.frx":0CAE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0D1A
               Keys            =   "frmMDWithhold.frx":0D38
               Spin            =   "frmMDWithhold.frx":0D82
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF3 
               Height          =   300
               Left            =   3135
               TabIndex        =   12
               Top             =   2250
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0DAA
               Caption         =   "frmMDWithhold.frx":0DCA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0E36
               Keys            =   "frmMDWithhold.frx":0E54
               Spin            =   "frmMDWithhold.frx":0E9E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA3 
               Height          =   300
               Left            =   5265
               TabIndex        =   13
               Top             =   2250
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0EC6
               Caption         =   "frmMDWithhold.frx":0EE6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":0F52
               Keys            =   "frmMDWithhold.frx":0F70
               Spin            =   "frmMDWithhold.frx":0FBA
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB4 
               Height          =   300
               Left            =   1005
               TabIndex        =   14
               Top             =   2640
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":0FE2
               Caption         =   "frmMDWithhold.frx":1002
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":106E
               Keys            =   "frmMDWithhold.frx":108C
               Spin            =   "frmMDWithhold.frx":10D6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF4 
               Height          =   300
               Left            =   3135
               TabIndex        =   15
               Top             =   2640
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":10FE
               Caption         =   "frmMDWithhold.frx":111E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":118A
               Keys            =   "frmMDWithhold.frx":11A8
               Spin            =   "frmMDWithhold.frx":11F2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA4 
               Height          =   300
               Left            =   5265
               TabIndex        =   16
               Top             =   2640
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":121A
               Caption         =   "frmMDWithhold.frx":123A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":12A6
               Keys            =   "frmMDWithhold.frx":12C4
               Spin            =   "frmMDWithhold.frx":130E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB5 
               Height          =   300
               Left            =   1005
               TabIndex        =   17
               Top             =   3030
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1336
               Caption         =   "frmMDWithhold.frx":1356
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":13C2
               Keys            =   "frmMDWithhold.frx":13E0
               Spin            =   "frmMDWithhold.frx":142A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF5 
               Height          =   300
               Left            =   3135
               TabIndex        =   18
               Top             =   3030
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1452
               Caption         =   "frmMDWithhold.frx":1472
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":14DE
               Keys            =   "frmMDWithhold.frx":14FC
               Spin            =   "frmMDWithhold.frx":1546
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA5 
               Height          =   300
               Left            =   5265
               TabIndex        =   19
               Top             =   3030
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":156E
               Caption         =   "frmMDWithhold.frx":158E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":15FA
               Keys            =   "frmMDWithhold.frx":1618
               Spin            =   "frmMDWithhold.frx":1662
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB6 
               Height          =   300
               Left            =   1005
               TabIndex        =   20
               Top             =   3420
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":168A
               Caption         =   "frmMDWithhold.frx":16AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1716
               Keys            =   "frmMDWithhold.frx":1734
               Spin            =   "frmMDWithhold.frx":177E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF6 
               Height          =   300
               Left            =   3135
               TabIndex        =   21
               Top             =   3420
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":17A6
               Caption         =   "frmMDWithhold.frx":17C6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1832
               Keys            =   "frmMDWithhold.frx":1850
               Spin            =   "frmMDWithhold.frx":189A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA6 
               Height          =   300
               Left            =   5265
               TabIndex        =   22
               Top             =   3420
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":18C2
               Caption         =   "frmMDWithhold.frx":18E2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":194E
               Keys            =   "frmMDWithhold.frx":196C
               Spin            =   "frmMDWithhold.frx":19B6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB7 
               Height          =   300
               Left            =   1005
               TabIndex        =   23
               Top             =   3810
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":19DE
               Caption         =   "frmMDWithhold.frx":19FE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1A6A
               Keys            =   "frmMDWithhold.frx":1A88
               Spin            =   "frmMDWithhold.frx":1AD2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF7 
               Height          =   300
               Left            =   3135
               TabIndex        =   24
               Top             =   3810
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1AFA
               Caption         =   "frmMDWithhold.frx":1B1A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1B86
               Keys            =   "frmMDWithhold.frx":1BA4
               Spin            =   "frmMDWithhold.frx":1BEE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA7 
               Height          =   300
               Left            =   5265
               TabIndex        =   25
               Top             =   3810
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1C16
               Caption         =   "frmMDWithhold.frx":1C36
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1CA2
               Keys            =   "frmMDWithhold.frx":1CC0
               Spin            =   "frmMDWithhold.frx":1D0A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB8 
               Height          =   300
               Left            =   1005
               TabIndex        =   26
               Top             =   4200
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1D32
               Caption         =   "frmMDWithhold.frx":1D52
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1DBE
               Keys            =   "frmMDWithhold.frx":1DDC
               Spin            =   "frmMDWithhold.frx":1E26
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF8 
               Height          =   300
               Left            =   3135
               TabIndex        =   27
               Top             =   4200
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1E4E
               Caption         =   "frmMDWithhold.frx":1E6E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1EDA
               Keys            =   "frmMDWithhold.frx":1EF8
               Spin            =   "frmMDWithhold.frx":1F42
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA8 
               Height          =   300
               Left            =   5265
               TabIndex        =   28
               Top             =   4200
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":1F6A
               Caption         =   "frmMDWithhold.frx":1F8A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":1FF6
               Keys            =   "frmMDWithhold.frx":2014
               Spin            =   "frmMDWithhold.frx":205E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB9 
               Height          =   300
               Left            =   1005
               TabIndex        =   29
               Top             =   4590
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":2086
               Caption         =   "frmMDWithhold.frx":20A6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":2112
               Keys            =   "frmMDWithhold.frx":2130
               Spin            =   "frmMDWithhold.frx":217A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF9 
               Height          =   300
               Left            =   3135
               TabIndex        =   30
               Top             =   4590
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":21A2
               Caption         =   "frmMDWithhold.frx":21C2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":222E
               Keys            =   "frmMDWithhold.frx":224C
               Spin            =   "frmMDWithhold.frx":2296
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA9 
               Height          =   300
               Left            =   5265
               TabIndex        =   31
               Top             =   4590
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":22BE
               Caption         =   "frmMDWithhold.frx":22DE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":234A
               Keys            =   "frmMDWithhold.frx":2368
               Spin            =   "frmMDWithhold.frx":23B2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtB10 
               Height          =   300
               Left            =   1005
               TabIndex        =   32
               Top             =   4980
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":23DA
               Caption         =   "frmMDWithhold.frx":23FA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":2466
               Keys            =   "frmMDWithhold.frx":2484
               Spin            =   "frmMDWithhold.frx":24CE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtF10 
               Height          =   300
               Left            =   3135
               TabIndex        =   33
               Top             =   4980
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":24F6
               Caption         =   "frmMDWithhold.frx":2516
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":2582
               Keys            =   "frmMDWithhold.frx":25A0
               Spin            =   "frmMDWithhold.frx":25EA
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtA10 
               Height          =   300
               Left            =   5265
               TabIndex        =   34
               Top             =   4980
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":2612
               Caption         =   "frmMDWithhold.frx":2632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":269E
               Keys            =   "frmMDWithhold.frx":26BC
               Spin            =   "frmMDWithhold.frx":2706
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber txtWTExemption 
               Height          =   300
               Left            =   1425
               TabIndex        =   3
               Top             =   630
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
               _ExtentY        =   529
               Calculator      =   "frmMDWithhold.frx":272E
               Caption         =   "frmMDWithhold.frx":274E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDWithhold.frx":27BA
               Keys            =   "frmMDWithhold.frx":27D8
               Spin            =   "frmMDWithhold.frx":2822
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,##0.00;;0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   1
               ForeColor       =   -2147483640
               Format          =   "###,###,##0.00"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ADD-ON"
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
               Left            =   5745
               TabIndex        =   66
               Top             =   1140
               Width           =   915
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "8"
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
               Left            =   -615
               TabIndex        =   65
               Top             =   4230
               Width           =   1470
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "7"
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
               Left            =   -375
               TabIndex        =   64
               Top             =   3855
               Width           =   1230
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "6"
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
               Left            =   -390
               TabIndex        =   63
               Top             =   3465
               Width           =   1230
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "5"
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
               Left            =   -390
               TabIndex        =   62
               Top             =   3075
               Width           =   1230
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "FACTOR"
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
               Left            =   3585
               TabIndex        =   61
               Top             =   1140
               Width           =   915
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "BRACKET"
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
               Left            =   1485
               TabIndex        =   60
               Top             =   1185
               Width           =   930
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "4"
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
               Left            =   -390
               TabIndex        =   59
               Top             =   2670
               Width           =   1230
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "2"
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
               Left            =   -720
               TabIndex        =   58
               Top             =   1890
               Width           =   1560
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "3"
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
               Left            =   -1005
               TabIndex        =   57
               Top             =   2280
               Width           =   1845
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "W/Tax Code"
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
               Left            =   -225
               TabIndex        =   56
               Top             =   315
               Width           =   1560
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "1"
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
               Left            =   0
               TabIndex        =   55
               Top             =   1500
               Width           =   855
            End
            Begin VB.Label Label18 
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
               Left            =   3045
               TabIndex        =   54
               Top             =   300
               Width           =   1560
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Exemption"
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
               Left            =   -240
               TabIndex        =   53
               Top             =   660
               Width           =   1560
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "9"
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
               Left            =   -600
               TabIndex        =   52
               Top             =   4635
               Width           =   1470
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "10"
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
               Left            =   -585
               TabIndex        =   51
               Top             =   5010
               Width           =   1470
            End
         End
      End
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   150
      TabIndex        =   67
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
         TabIndex        =   71
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
         TabIndex        =   72
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
         TabIndex        =   73
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
      TabIndex        =   74
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
Attribute VB_Name = "frmMDWithhold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Recordsets
Option Explicit
Dim WT As ADODB.Recordset

'Booleans
Dim mAdd As Boolean
Dim mEdit As Boolean
Dim mTransActive As Boolean

'storage
Dim mCode As Integer
Dim mWTSortField As String

Private Sub Form_Load()

    Add_MDIButton Me.Name, TitleBar.Caption
    
    Me.Width = 8940
    Me.Height = 7170
    FormCenter Me
    Call LoadWT
    
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

Private Sub Edit_Record()
If cmdMenu(1).Caption = "&Edit" Then
    mTransActive = True
    cmdMenu(1).Caption = "&Save"
    Lock_Button "FTFTFF", cmdMenu, 5
    frmeWT.Enabled = True
    gridWT.Enabled = False
    mCode = txtWTCode.Text
    tabWT.CurrTab = 0
    SafeSetFocus txtWTExemption
    mEdit = True
Else
    If Not IsNumeric(txtWTExemption.Text) Then
        MsgBox "Exemption field must be filled up.", vbInformation
        SafeSetFocus txtWTExemption
        Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    
    ConMain.BeginTrans
    
        mTransActive = True
        ConMain.Execute "update WT set description = '" & txtWTDescription.Text & "', exemption = " & txtWTExemption.Value & ", " & _
            "b1 =" & txtB1.Value & ", f1 =" & txtF1.Value & ", a1 = " & txtA1.Value & ", " & _
            "b2 =" & txtB2.Value & ", f2 =" & txtF2.Value & ", a2 =" & txtA2.Value & ", " & _
            "b3 =" & txtB3.Value & ", f3 =" & txtF3.Value & ", a3 =" & txtA3.Value & ", " & _
            "b4 = " & txtB4.Value & ", f4 =" & txtF4.Value & ", a4 =" & txtA4.Value & ", " & _
            "b5 =" & txtB5.Value & ", f5 = " & txtF5.Value & ", a5 =" & txtA5.Value & ", " & _
            "b6 =" & txtB6.Value & ", f6 =" & txtF6.Value & ", a6 =" & txtA6.Value & ", " & _
            "b7 =" & txtB7.Value & ", f7 =" & txtF7.Value & ", a7 =" & txtA7.Value & ", " & _
            "b8 =" & txtB8.Value & ", f8 =" & txtF8.Value & ", a8 =" & txtA8.Value & ", " & _
            "b9 =" & txtB9.Value & ", f9 =" & txtF9.Value & ", a9 =" & txtA9.Value & ", " & _
            "b10 =" & txtB10.Value & ", f10 =" & txtF10.Value & ",a10 = " & txtA10.Value & " where wtcode = '" & txtWTCode.Text & "'"
    
    ConMain.CommitTrans
    
    gridWT.Enabled = True
    frmeWT.Enabled = False
    WT.Requery
    pointmetdg gridWT, WT, "wtcode", mCode
    mEdit = False
    mTransActive = False
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
    tabWT.CurrTab = 1
    
End If
End Sub

Private Sub Delete_Record()
If WT.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        gridWT.Delete
    End If
End If
End Sub

Private Sub Cancel_Transaction()
If mAdd = True Then
    cmdMenu(0).Caption = "&New"
    If WT.RecordCount > 0 Then
        Lock_Button "TTTFTT", cmdMenu, 5
    Else
        Lock_Button "TFFFTT", cmdMenu, 5
    End If

    mAdd = False
End If
If mEdit = True Then
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
    mEdit = False
End If
frmeWT.Enabled = False
gridWT.Enabled = True
gridWT_RowColChange gridWT.Row, gridWT.Col
tabWT.CurrTab = 1
End Sub

Private Sub Print_Record()

End Sub

Private Sub Close_Form()
Unload Me
End Sub

Private Sub ClearFields()
        txtWTCode.Text = "AUTO GENERATED..."
        txtWTDescription.Text = ""
        txtWTExemption.Value = 0
        txtB1.Value = 0
        txtB2.Value = 0
        txtB3.Value = 0
        txtB4.Value = 0
        txtB5.Value = 0
        txtB6.Value = 0
        txtB7.Value = 0
        txtB8.Value = 0
        txtB9.Value = 0
        txtB10.Value = 0
        
        txtF1.Value = 0
        txtF2.Value = 0
        txtF3.Value = 0
        txtF4.Value = 0
        txtF5.Value = 0
        txtF6.Value = 0
        txtF7.Value = 0
        txtF8.Value = 0
        txtF9.Value = 0
        txtF10.Value = 0
        
        txtA1.Value = 0
        txtA2.Value = 0
        txtA3.Value = 0
        txtA4.Value = 0
        txtA5.Value = 0
        txtA6.Value = 0
        txtA7.Value = 0
        txtA8.Value = 0
        txtA9.Value = 0
        txtA10.Value = 0
End Sub

Private Sub Add_Record()
If cmdMenu(0).Caption = "&New" Then
    mTransActive = True
    cmdMenu(0).Caption = "&Save"
    Lock_Button "TFFTFF", cmdMenu, 5
    frmeWT.Enabled = True
    gridWT.Enabled = False
    Call ClearFields
    tabWT.CurrTab = 0
    SafeSetFocus txtWTExemption
    mAdd = True
Else
    If Not IsNumeric(txtWTExemption.Text) Then
        MsgBox "Exemption field must be filled up.", vbInformation
        SafeSetFocus txtWTExemption
        Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        txtWTCode.Text = LastCode("WT")
        ConMain.Execute "insert into wt values ('" & txtWTCode.Text & "', '" & txtWTDescription.Text & "', " & txtWTExemption.Value & ", " & _
            "" & txtB1.Value & ", " & txtF1.Value & ", " & txtA1.Value & ", " & _
            "" & txtB2.Value & ", " & txtF2.Value & ", " & txtA2.Value & ", " & _
            "" & txtB3.Value & ", " & txtF3.Value & ", " & txtA3.Value & ", " & _
            "" & txtB4.Value & ", " & txtF4.Value & ", " & txtA4.Value & ", " & _
            "" & txtB5.Value & ", " & txtF5.Value & ", " & txtA5.Value & ", " & _
            "" & txtB6.Value & ", " & txtF6.Value & ", " & txtA6.Value & ", " & _
            "" & txtB7.Value & ", " & txtF7.Value & ", " & txtA7.Value & ", " & _
            "" & txtB8.Value & ", " & txtF8.Value & ", " & txtA8.Value & ", " & _
            "" & txtB9.Value & ", " & txtF9.Value & ", " & txtA9.Value & ", " & _
            "" & txtB10.Value & ", " & txtF10.Value & ", " & txtA10.Value & ")"
    ConMain.CommitTrans
    
    gridWT.Enabled = True
    frmeWT.Enabled = False
    mCode = txtWTCode.Text
    WT.Requery
    pointmetdg gridWT, WT, "wtcode", mCode
    mAdd = False
    mTransActive = False
    cmdMenu(0).Caption = "&New"
    Lock_Button "TTTFTT", cmdMenu, 5
    tabWT.CurrTab = 1
    
End If

End Sub

Private Sub LoadWT()
DoEvents
NetOpen WT, "select * from WT order by DESCRIPTION"
DoEvents
If WT.State = adStateOpen Then
    If WT.RecordCount > 0 Then
        WT.MoveFirst
        Lock_Button "TTTFTT", cmdMenu, 5
    Else
        Lock_Button "TFFFTT", cmdMenu, 5
    End If
    Set gridWT.DataSource = WT
    mWTSortField = "WTcode"
End If
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
    
    With tabWT
        .Top = fraButton.Top + fraButton.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
    
    With gridWT
        .Left = 150
        .Width = Me.ScaleWidth - 300
        .Height = Me.ScaleHeight - .Top - 2000
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMDWithhold = Nothing
End Sub

Sub FormCenter(Frm As Form)
    Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub

Private Sub gridWT_HeadClick(ByVal ColIndex As Integer)
If WT.RecordCount > 0 Then
    mWTSortField = gridWT.Columns(ColIndex).DataField
    WT.Sort = mWTSortField
End If
End Sub

Private Sub gridWT_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With WT
    If .RecordCount > 0 Then
        txtWTCode.Text = !wtcode
        txtWTDescription.Text = !Description
        txtWTExemption.Value = !exemption
        txtB1.Value = !b1
        txtB2.Value = !b2
        txtB3.Value = !b3
        txtB4.Value = !b4
        txtB5.Value = !b5
        txtB6.Value = !b6
        txtB7.Value = !b7
        txtB8.Value = !b8
        txtB9.Value = !b9
        txtB10.Value = !b10
        
        txtF1.Value = !f1
        txtF2.Value = !f2
        txtF3.Value = !f3
        txtF4.Value = !f4
        txtF5.Value = !f5
        txtF6.Value = !f6
        txtF7.Value = !f7
        txtF8.Value = !f8
        txtF9.Value = !f9
        txtF10.Value = !f10
        
        txtA1.Value = !a1
        txtA2.Value = !a2
        txtA3.Value = !a3
        txtA4.Value = !a4
        txtA5.Value = !a5
        txtA6.Value = !a6
        txtA7.Value = !a7
        txtA8.Value = !a8
        txtA9.Value = !a9
        txtA10.Value = !a10
    Else
        Call ClearFields
    End If
End With
End Sub

Private Sub tabWT_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
If mAdd = True Then
    Cancel = 1
End If
If mEdit = True Then
    Cancel = 1
End If
End Sub



Private Sub txtB1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtB10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtF10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtA10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtwtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub txtWTDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
