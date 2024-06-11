VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDHoliday 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   7080
   ClientLeft      =   2850
   ClientTop       =   4665
   ClientWidth     =   8910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMDHoliday.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab tabHoliday 
      Height          =   6405
      Left            =   105
      TabIndex        =   11
      Top             =   570
      Width           =   7590
      _cx             =   13388
      _cy             =   11298
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
      Caption         =   "Maintain Holidays|View Holidays"
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
      Begin C1SizerLibCtl.C1Elastic SizerMAType 
         Height          =   6090
         Left            =   15
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   7560
         _cx             =   13335
         _cy             =   10742
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
         Begin TrueOleDBGrid80.TDBDropDown dcboBranches 
            Height          =   1365
            Left            =   1665
            TabIndex        =   34
            Top             =   3735
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2408
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Branch Code"
            Columns(0).DataField=   "BranchCode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Branch Name"
            Columns(1).DataField=   "Branch"
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
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
            EmptyRows       =   0   'False
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
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H8000000D&"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HDAFAEF&"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0"
            _StyleDefs(40)  =   "Named:id=34:Heading"
            _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=34,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=35:Footing"
            _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=36:Selected"
            _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=37:Caption"
            _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(49)  =   "Named:id=38:HighlightRow"
            _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(51)  =   "Named:id=39:EvenRow"
            _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=40:OddRow"
            _StyleDefs(54)  =   ":id=40,.parent=33"
            _StyleDefs(55)  =   "Named:id=41:RecordSelector"
            _StyleDefs(56)  =   ":id=41,.parent=34"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText txtBranch 
            Height          =   255
            Left            =   330
            TabIndex        =   33
            Top             =   3150
            Visible         =   0   'False
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   450
            Caption         =   "frmMDHoliday.frx":6852
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDHoliday.frx":68B8
            Key             =   "frmMDHoliday.frx":68D6
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
         Begin VB.Frame frmeHoliday 
            BackColor       =   &H00F6F8F8&
            Enabled         =   0   'False
            Height          =   2205
            Left            =   120
            TabIndex        =   26
            Top             =   75
            Width           =   7290
            Begin VB.Frame frmeSpecialRegular 
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   0  'None
               Height          =   795
               Left            =   3855
               TabIndex        =   32
               Top             =   150
               Width           =   3075
               Begin VB.OptionButton optSpecial 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Special Holiday"
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   165
                  TabIndex        =   5
                  Top             =   480
                  Width           =   1710
               End
               Begin VB.OptionButton optRegular 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Legal Holiday"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   165
                  TabIndex        =   4
                  Top             =   195
                  Width           =   2100
               End
            End
            Begin TDBText6Ctl.TDBText txtYear 
               Height          =   300
               Left            =   1725
               TabIndex        =   2
               Top             =   240
               Width           =   2010
               _Version        =   65536
               _ExtentX        =   3545
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":691A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6980
               Key             =   "frmMDHoliday.frx":699E
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
               AlignVertical   =   0
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
            Begin VB.OptionButton optAllBranches 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "All Branches"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   1710
               TabIndex        =   8
               Top             =   1800
               Width           =   1620
            End
            Begin VB.OptionButton optSelectedBranches 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Selected Branches"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   3495
               TabIndex        =   9
               Top             =   1800
               Width           =   2130
            End
            Begin TDBText6Ctl.TDBText txtDescription 
               Height          =   300
               Left            =   1725
               TabIndex        =   7
               Top             =   1410
               Width           =   5220
               _Version        =   65536
               _ExtentX        =   9208
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":69E2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6A4E
               Key             =   "frmMDHoliday.frx":6A6C
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
            Begin TDBDate6Ctl.TDBDate txtHolidayDate 
               Height          =   300
               Left            =   1725
               TabIndex        =   3
               Top             =   645
               Width           =   2010
               _Version        =   65536
               _ExtentX        =   3545
               _ExtentY        =   529
               Calendar        =   "frmMDHoliday.frx":6AB0
               Caption         =   "frmMDHoliday.frx":6BB6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6C1C
               Keys            =   "frmMDHoliday.frx":6C3A
               Spin            =   "frmMDHoliday.frx":6C98
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
               MinDate         =   2
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   39464
               CenturyMode     =   0
            End
            Begin TDBText6Ctl.TDBText txtHolidayName 
               Height          =   300
               Left            =   1725
               TabIndex        =   6
               Top             =   1020
               Width           =   5220
               _Version        =   65536
               _ExtentX        =   9208
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":6CC0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6D2C
               Key             =   "frmMDHoliday.frx":6D4A
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
            Begin VB.Label Label1 
               BackColor       =   &H00F6F8F8&
               Caption         =   "Year"
               Height          =   300
               Left            =   1215
               TabIndex        =   31
               Top             =   285
               Width           =   675
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Apply to"
               Height          =   255
               Left            =   780
               TabIndex        =   30
               Top             =   1785
               Width           =   855
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Holiday Name"
               Height          =   255
               Left            =   150
               TabIndex        =   29
               Top             =   1050
               Width           =   1470
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Holiday Date"
               Height          =   255
               Left            =   45
               TabIndex        =   28
               Top             =   675
               Width           =   1560
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               Height          =   255
               Left            =   390
               TabIndex        =   27
               Top             =   1440
               Width           =   1230
            End
         End
         Begin TrueOleDBGrid80.TDBGrid gridBranches 
            Height          =   3570
            Left            =   120
            TabIndex        =   10
            Top             =   2385
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   6297
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Branch Code"
            Columns(0).DataField=   "BranchCode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).ValueItems(0)._DefaultItem=   0
            Columns(1).ValueItems(0).Value=   ""
            Columns(1).ValueItems(0).Value.vt=   8
            Columns(1).ValueItems(0).DisplayValue=   "1"
            Columns(1).ValueItems(0).DisplayValue.vt=   8
            Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(1).ValueItems.Count=   1
            Columns(1).Caption=   "Branch"
            Columns(1).DataField=   "Branch"
            Columns(1).DropDown=   "dcboBranches"
            Columns(1).DropDown.vt=   8
            Columns(1).ExternalEditor=   "txtBranch"
            Columns(1).ExternalEditor.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   84
            Columns(2)._MaxComboItems=   5
            Columns(2).ValueItems(0)._DefaultItem=   0
            Columns(2).ValueItems(0).Value=   "1"
            Columns(2).ValueItems(0).Value.vt=   8
            Columns(2).ValueItems(0).DisplayValue=   "1"
            Columns(2).ValueItems(0).DisplayValue.vt=   8
            Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems(1)._DefaultItem=   0
            Columns(2).ValueItems(1).Value=   "2"
            Columns(2).ValueItems(1).Value.vt=   8
            Columns(2).ValueItems(1).DisplayValue=   "0"
            Columns(2).ValueItems(1).DisplayValue.vt=   8
            Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems.Count=   2
            Columns(2).Caption=   "Included"
            Columns(2).DataField=   "Included"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7567"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7488"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(1).AutoDropDown=1"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=3625"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3545"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
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
            EditDropDown    =   0   'False
            Enabled         =   0   'False
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Named:id=33:Normal"
            _StyleDefs(47)  =   ":id=33,.parent=0"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   ":id=34,.wraptext=-1"
            _StyleDefs(51)  =   "Named:id=35:Footing"
            _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   "Named:id=36:Selected"
            _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=37:Caption"
            _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(57)  =   "Named:id=38:HighlightRow"
            _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(59)  =   "Named:id=39:EvenRow"
            _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(61)  =   "Named:id=40:OddRow"
            _StyleDefs(62)  =   ":id=40,.parent=33"
            _StyleDefs(63)  =   "Named:id=41:RecordSelector"
            _StyleDefs(64)  =   ":id=41,.parent=34"
            _StyleDefs(65)  =   "Named:id=42:FilterBar"
            _StyleDefs(66)  =   ":id=42,.parent=33"
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   6090
         Left            =   8205
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   7560
         _cx             =   13335
         _cy             =   10742
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
         Begin TDBText6Ctl.TDBText txtSearchBoxHoliday 
            Height          =   300
            Left            =   1530
            TabIndex        =   0
            Top             =   150
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
            _ExtentY        =   529
            Caption         =   "frmMDHoliday.frx":6D8E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDHoliday.frx":6DFA
            Key             =   "frmMDHoliday.frx":6E18
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
         Begin TrueOleDBGrid80.TDBGrid gridHolidays 
            Height          =   5400
            Left            =   105
            TabIndex        =   1
            Top             =   570
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   9525
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Year"
            Columns(0).DataField=   "curryear"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Holiday Date"
            Columns(1).DataField=   "HolidayDate"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Holiday"
            Columns(2).DataField=   "Holiday"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Holiday Type"
            Columns(3).DataField=   "Description"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Regular"
            Columns(4).DataField=   "Regular"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Special"
            Columns(5).DataField=   "Special"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Applicability"
            Columns(6).DataField=   "Applicability"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2884"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2805"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3598"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3519"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(62)  =   "Named:id=33:Normal"
            _StyleDefs(63)  =   ":id=33,.parent=0"
            _StyleDefs(64)  =   "Named:id=34:Heading"
            _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   ":id=34,.wraptext=-1"
            _StyleDefs(67)  =   "Named:id=35:Footing"
            _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(69)  =   "Named:id=36:Selected"
            _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=37:Caption"
            _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(73)  =   "Named:id=38:HighlightRow"
            _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(75)  =   "Named:id=39:EvenRow"
            _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(77)  =   "Named:id=40:OddRow"
            _StyleDefs(78)  =   ":id=40,.parent=33"
            _StyleDefs(79)  =   "Named:id=41:RecordSelector"
            _StyleDefs(80)  =   ":id=41,.parent=34"
            _StyleDefs(81)  =   "Named:id=42:FilterBar"
            _StyleDefs(82)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   195
            Width           =   915
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   6090
         Left            =   8505
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Width           =   7560
         _cx             =   13335
         _cy             =   10742
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
            TabIndex        =   16
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   17
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":6E5C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6EC8
               Key             =   "frmMDHoliday.frx":6EE6
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
               TabIndex        =   18
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":6F2A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":6F96
               Key             =   "frmMDHoliday.frx":6FB4
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
               TabIndex        =   19
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDHoliday.frx":6FF8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDHoliday.frx":7064
               Key             =   "frmMDHoliday.frx":7082
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
               TabIndex        =   22
               Top             =   945
               Width           =   990
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
               TabIndex        =   21
               Top             =   300
               Width           =   1335
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
               TabIndex        =   20
               Top             =   615
               Width           =   1635
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   23
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "frmMDHoliday.frx":70C6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDHoliday.frx":7132
            Key             =   "frmMDHoliday.frx":7150
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
            TabIndex        =   24
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
            FormatString    =   $"frmMDHoliday.frx":7194
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
            TabIndex        =   25
            Top             =   240
            Width           =   915
         End
      End
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
      TabIndex        =   42
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
Attribute VB_Name = "frmMDHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'recordsets
Dim branch As ADODB.Recordset
Dim Holiday As ADODB.Recordset

'temporary recordsets
Dim TempDB As ADODB.Recordset

'Booleans
Dim mAdd As Boolean
Dim mEdit As Boolean
Dim mTransActive As Boolean
Dim mDupError As Boolean
Dim mLoadComplete As Boolean

'storage
Dim mCode As Integer
Dim mYear As String
Dim mHolidaySortField As String

Private Sub LoadHoliday()

    DoEvents
    NetOpen Holiday, "select a.currYear as curryear, a.Holidaydate as Holidaydate, " & _
                         "a.Holiday as Holiday, a.description as Description, " & _
                         "case when a.regular = 1 then 'Yes' else 'No' end as Regular, " & _
                         "case when a.special = 1 then 'Yes' else 'No' end as Special, " & _
                         "case when a.applicability = 1 then 'All Branches' else 'Selected Branches' end as Applicability " & _
                         "From Holiday a order by holidaydate"
    DoEvents
    If Holiday.State = adStateOpen Then
        If Holiday.RecordCount > 0 Then
            Holiday.MoveFirst
            Lock_Button "TTTFTT", cmdMenu, 5
        Else
            Lock_Button "TFFFTT", cmdMenu, 5
        End If
        Set gridHolidays.DataSource = Holiday
        mHolidaySortField = "Holiday"
    End If
    
End Sub

Private Sub InitTempDB()
  Set TempDB = New ADODB.Recordset
  With TempDB
    .Fields.Append "BranchCode", adChar, 7, adFldIsNullable
    .Fields.Append "Branch", adChar, 100, adFldIsNullable
    .Fields.Append "Included", adInteger, 1, adFldIsNullable
    .Open
  End With
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
Call InitTempDB
Set gridBranches.DataSource = TempDB
'fill the branch list with branches
txtYear.Text = Format(Date, "YYYY")
txtHolidayDate.Text = ""
txtHolidayName.Text = ""
txtDescription.Text = ""
optSelectedBranches.Value = True
optRegular.Value = True
End Sub

Private Sub Add_Record()
If cmdMenu(0).Caption = "&New" Then
    tabHoliday.CurrTab = 0
    mAdd = True
    mTransActive = True
    cmdMenu(0).Caption = "&Save"
    frmeHoliday.Enabled = True
    Lock_Button "TFFTFF", cmdMenu, 5
    Call ClearFields
    SafeSetFocus txtHolidayDate
Else
    If IsDate(txtHolidayDate.Text) = False Then
        MsgBox "You have to add the date of the holiday specified.", vbInformation
        SafeSetFocus txtHolidayDate
        Exit Sub
    End If
    If Trim(txtHolidayName.Text) = "" Then
        MsgBox "Name of Holiday must be specified.", vbInformation
        SafeSetFocus txtHolidayName
        Exit Sub
    End If
    If TempDB.RecordCount = 0 Then
        MsgBox "This holiday information must be assigned to at lease one branch.", vbInformation
        optAllBranches.Value = True
        Exit Sub
    End If
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        'save holiday info
        ConMain.Execute "insert into Holiday values ('" & txtYear.Text & "', '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "', '" & CSQ(txtHolidayName.Text) & "', '" & CSQ(txtDescription.Text) & "', '" & IIf(optRegular.Value = True, 1, 0) & "', '" & IIf(optSpecial.Value = True, 1, 0) & "', '" & IIf(optAllBranches.Value = True, 1, 0) & "')"
        'save branches
        If TempDB.RecordCount > 0 Then
            TempDB.MoveFirst
            Do While Not TempDB.EOF
                ConMain.Execute "insert into holidaybranchinclude values('" & txtYear.Text & "', '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "', '" & TempDB!branchcode & "', " & IIf(Val(TempDB!included) > 0, Val(TempDB!included), Val(TempDB!included) * -1) & ")"
                TempDB.MoveNext
            Loop
        End If
    ConMain.CommitTrans
    gridHolidays.Enabled = True
    frmeHoliday.Enabled = False
    txtSearchBoxHoliday.Enabled = True
    mCode = txtHolidayDate.Text
    Holiday.Requery
    pointmetdg gridHolidays, Holiday, "Holidaydate", mCode
    mAdd = False
    mTransActive = False
    cmdMenu(0).Caption = "&New"
    Lock_Button "TTTFTT", cmdMenu, 5
    tabHoliday.CurrTab = 1
End If
End Sub

Private Sub Edit_Record()
If Holiday.RecordCount > 0 Then
    If cmdMenu(1).Caption = "&Edit" Then
        mCode = Format(txtHolidayDate.Text, "YYYY-MM-DD")
        mYear = txtYear.Text
        tabHoliday.CurrTab = 0
        mEdit = True
        mTransActive = True
        frmeHoliday.Enabled = True
        gridBranches.Enabled = True
        cmdMenu(1).Caption = "&Save"
        Lock_Button "FTFTFF", cmdMenu, 5
        SafeSetFocus txtHolidayDate
    Else
        If IsDate(txtHolidayDate.Text) = False Then
            MsgBox "You have to add the date of the holiday specified.", vbInformation
            SafeSetFocus txtHolidayDate
            Exit Sub
        End If
        If Trim(txtHolidayName.Text) = "" Then
            MsgBox "Name of Holiday must be specified.", vbInformation
            SafeSetFocus txtHolidayName
            Exit Sub
        End If
        If TempDB.RecordCount = 0 Then
            MsgBox "This holiday information must be assigned to at lease one branch.", vbInformation
            optAllBranches.Value = True
            Exit Sub
        End If
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
            mTransActive = True
            'save holiday info
            ConMain.Execute "update Holiday set curryear = '" & txtYear.Text & "', " & _
                                  "holidaydate = '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "', " & _
                                  "description = '" & CSQ(txtDescription.Text) & "', " & _
                                  "regular = " & IIf(optRegular.Value = True, 1, 0) & ", " & _
                                  "special = " & IIf(optSpecial.Value = True, 1, 0) & ", " & _
                                  "applicability = " & IIf(optAllBranches.Value = True, 1, 0) & "  where curryear = '" & mYear & "' and holidaydate = '" & mCode & "'"
            'save branches
            If TempDB.RecordCount > 0 Then
                ConMain.Execute "delete from holidaybranchinclude where curryear = '" & txtYear.Text & "' and holidaydate = '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "'"
            End If
            If TempDB.RecordCount > 0 Then
                TempDB.MoveFirst
                Do While Not TempDB.EOF
                    ConMain.Execute "insert into holidaybranchinclude values('" & txtYear.Text & "', '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "', '" & TempDB!branchcode & "', " & IIf(Val(TempDB!included) > 0, Val(TempDB!included), Val(TempDB!included) * -1) & ")"
                    TempDB.MoveNext
                Loop
            End If
        ConMain.CommitTrans
        gridHolidays.Enabled = True
        frmeHoliday.Enabled = False
        txtSearchBoxHoliday.Enabled = True
        Holiday.Requery
        pointmetdg gridHolidays, Holiday, "Holidaydate", txtHolidayDate.Text
        mEdit = False
        mTransActive = False
        cmdMenu(1).Caption = "&Edit"
        Lock_Button "TTTFTT", cmdMenu, 5
        tabHoliday.CurrTab = 1
    End If
End If
End Sub

Private Sub Delete_Record()
If Holiday.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        ConMain.Execute "delete from holiday where curryear = '" & txtYear.Text & "' and holidaydate = '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "'"
        ConMain.Execute "delete from holidaybranchinclude where curryear = '" & txtYear.Text & "' and holidaydate = '" & Format(txtHolidayDate.Text, "YYYY-MM-DD") & "'"
        Holiday.Requery
    End If
    tabHoliday.CurrTab = 1
End If
End Sub

Private Sub Cancel_Transaction()
If mAdd = True Then
    cmdMenu(0).Caption = "&New"
    Lock_Button "TTTFTT", cmdMenu, 5
    mAdd = False
End If
If mEdit = True Then
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
    mEdit = False
End If
frmeHoliday.Enabled = False
txtSearchBoxHoliday.Enabled = True
gridBranches.Enabled = True
gridHolidays.Enabled = True
gridHolidays_RowColChange gridHolidays.Row, gridHolidays.Col
tabHoliday.CurrTab = 1
End Sub

Private Sub Print_Record()

End Sub

Private Sub Close_Form()
Unload Me
End Sub

Private Sub dcboBranches_DropDownOpen()
    dcboBranches.Width = gridBranches.Columns(1).Width
End Sub

Private Sub dcboBranches_RowChange()
If Trim(dcboBranches.Columns(1).Value) <> "" Then
    gridBranches.Columns(0).Value = Trim(dcboBranches.Columns(0).Value)
    txtBranch.Text = Trim(dcboBranches.Columns(1).Value)
Else
    gridBranches.Columns(0).Value = ""
End If
End Sub

Private Sub Form_Activate()
'bind dropdown
BindDropDown dcboBranches, "select branchcode, branch from branch", "branch"
End Sub

Private Sub Form_Load()
    
   Add_MDIButton Me.Name, TitleBar.Caption
   
    'bind dropdown
    BindDropDown dcboBranches, "select branchcode, branch from branch", "branch"
    Call LoadHoliday
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()
On Error Resume Next

    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With tabHoliday
        .Top = frabutton.Top + frabutton.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMDHoliday = Nothing
End Sub

Sub FormCenter(Frm As Form)
    Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub


Private Sub gridBranches_AfterColUpdate(ByVal ColIndex As Integer)
If TempDB.RecordCount > 0 Then
    If IsEmpty(gridBranches.Columns(2).Value) = False Then
        If gridBranches.Columns(2).Value = 2 Then
            gridBranches.Delete
        End If
    End If
End If
End Sub

Private Sub gridBranches_BeforeRowColChange(Cancel As Integer)
If IsNull(gridBranches.Columns(1).Value) = False Then
    If DuplicateCheck(gridBranches, gridBranches.Columns(1).Value, 1) = True Then
        MsgBox "Branch already exist on the list.", vbInformation
        Cancel = 1
        SafeSetFocus gridBranches
        SendKeys " "
        mDupError = True
    End If
End If
If Trim(gridBranches.Columns(1).Value) = "" Then
    Cancel = 1
    SafeSetFocus gridBranches
    SendKeys " "
End If
End Sub

Private Sub gridBranches_BeforeUpdate(Cancel As Integer)
If Trim(gridBranches.Columns(1).Value) = "" Then
    Cancel = 1
End If
End Sub

Private Sub gridBranches_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If mDupError = True Then
        gridBranches.Delete
        mDupError = False
    End If
End If
End Sub

Private Sub optAllBranches_Click()
If Me.Visible = True Then
    If mLoadComplete = False Then
        Call InitTempDB
        Set gridBranches.DataSource = TempDB
        NetOpen branch, "select branchcode, branch, 1 as include from branch order by branch"
        If branch.RecordCount > 0 Then
            branch.MoveFirst
            Do While Not branch.EOF
                With TempDB
                    .AddNew
                    .Fields("branchcode") = branch!branchcode
                    .Fields("branch") = branch!branch
                    .Fields("Included") = branch!include
                    .Update
                End With
            branch.MoveNext
            Loop
        End If
    Else
        mLoadComplete = False
    End If
End If
gridBranches.Enabled = False
End Sub

Private Sub optSelectedBranches_Click()
If Me.Visible = True Then
    gridBranches.Enabled = True
End If
End Sub

Private Sub tabHoliday_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
If mAdd = True Then
    Cancel = 1
    SafeSetFocus txtHolidayDate
End If
If mEdit = True Then
    Cancel = 1
    SafeSetFocus txtHolidayDate
End If
If mAdd = False And mEdit = False Then
    gridBranches.Enabled = False
End If
End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
    SearchRecord KeyAscii, txtBranch, dcboBranches.DataSource, txtBranch.Text, "branch"
End Sub

Private Sub txtBranch_LostFocus()
SafeSetFocus gridBranches
End Sub


Private Sub gridHolidays_HeadClick(ByVal ColIndex As Integer)
If Holiday.RecordCount > 0 Then
    mHolidaySortField = gridHolidays.Columns(ColIndex).DataField
    Holiday.Sort = mHolidaySortField
End If
End Sub

Private Sub gridHolidays_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With Holiday
    If .RecordCount > 0 Then
        txtYear.Text = !curryear
        txtHolidayDate.Text = Format(!holidaydate, "MM/DD/YYYY")
        txtHolidayName.Text = !Holiday
        txtDescription.Text = !Description
        optRegular.Value = IIf(!regular = "Yes", True, False)
        optSpecial.Value = IIf(!special = "Yes", True, False)
        optAllBranches.Value = IIf(!applicability = "All Branches", True, False)
        optSelectedBranches.Value = IIf(!applicability = "Selected Branches", True, False)
    
        Call InitTempDB
        NetOpen branch, "select a.branchcode as branchcode, b.branch as branch, a.include as include from holidaybranchinclude a inner join branch b on a.branchcode = b.branchcode and  a.curryear = '" & !curryear & "' and a.holidaydate = '" & Format(!holidaydate, "YYYY-MM-DD") & "'"
        If branch.RecordCount > 0 Then
            branch.MoveFirst
            Do While Not branch.EOF
                With TempDB
                    .AddNew
                    .Fields("branchcode") = branch!branchcode
                    .Fields("branch") = branch!branch
                    .Fields("included") = branch!include
                    .Update
                End With
                branch.MoveNext
            Loop
        End If
        If TempDB.State = adStateOpen Then
            Set gridBranches.DataSource = TempDB
        End If
        mLoadComplete = False
    Else
        txtYear.Text = ""
        txtHolidayDate.Text = ""
        txtHolidayName.Text = ""
        txtDescription.Text = ""
        optRegular.Value = False
        optSpecial.Value = False
        optAllBranches.Value = False
        optSelectedBranches.Value = False
        Call InitTempDB
        If TempDB.State = adStateOpen Then
            Set gridBranches.DataSource = TempDB
        End If
        mLoadComplete = False
    End If
End With
End Sub

Private Sub txtSearchBoxHoliday_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
Else
    SearchRecord KeyAscii, txtSearchBoxHoliday, Holiday, txtSearchBoxHoliday.Text, mHolidaySortField
End If
End Sub
