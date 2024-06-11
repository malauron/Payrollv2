VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDPayrollPeriod 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   12975
   Tag             =   "Payroll Period"
   Begin VB.PictureBox pic1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12975
      TabIndex        =   27
      Top             =   0
      Width           =   12975
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
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
         Index           =   5
         Left            =   135
         TabIndex        =   28
         Top             =   225
         Width           =   5445
      End
   End
   Begin VB.Frame fraHolder 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "mark"
      Height          =   7770
      Left            =   270
      TabIndex        =   18
      Top             =   840
      Width           =   12945
      Begin VB.Frame fraButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   30
         TabIndex        =   22
         Top             =   7185
         Width           =   12315
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   0
            Left            =   15
            TabIndex        =   23
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
            Image           =   "frmMDPayrollPeriod.frx":0000
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   1
            Left            =   1470
            TabIndex        =   24
            Top             =   45
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   820
            Caption         =   "&SAVE"
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
            Image           =   "frmMDPayrollPeriod.frx":1CDA
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   2
            Left            =   6030
            TabIndex        =   25
            Top             =   45
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
            Image           =   "frmMDPayrollPeriod.frx":2454
            ImgSize         =   24
            cBack           =   14737632
         End
         Begin lvButton.lvButtons_H cmdMenu 
            Height          =   465
            Index           =   3
            Left            =   2925
            TabIndex        =   26
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
            Image           =   "frmMDPayrollPeriod.frx":412E
            cBack           =   14737632
         End
      End
      Begin C1SizerLibCtl.C1Tab tabEmployee 
         Height          =   6915
         Left            =   45
         TabIndex        =   19
         Top             =   255
         Width           =   12900
         _cx             =   22754
         _cy             =   12197
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         BackColor       =   16777215
         ForeColor       =   4210752
         FrontTabColor   =   14737632
         BackTabColor    =   14737632
         TabOutlineColor =   12632256
         FrontTabForeColor=   4210752
         Caption         =   "Details"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   4
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   1
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6600
            Left            =   15
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   300
            Width           =   12870
            _cx             =   22701
            _cy             =   11642
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14737632
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
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   6600
               Index           =   0
               Left            =   30
               TabIndex        =   21
               Top             =   -45
               Width           =   12810
               Begin VB.Frame Frame3 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Search Payroll Period"
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
                  Height          =   1020
                  Left            =   60
                  TabIndex        =   35
                  Top             =   120
                  Width           =   6780
                  Begin TDBText6Ctl.TDBText txtSearch 
                     Height          =   300
                     Left            =   1215
                     TabIndex        =   1
                     Top             =   600
                     Width           =   5115
                     _Version        =   65536
                     _ExtentX        =   9022
                     _ExtentY        =   529
                     Caption         =   "frmMDPayrollPeriod.frx":4A08
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":4A74
                     Key             =   "frmMDPayrollPeriod.frx":4A92
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
                  Begin lvButton.lvButtons_H cmdSearch 
                     Height          =   315
                     Left            =   6360
                     TabIndex        =   36
                     Top             =   600
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
                  Begin TrueOleDBList80.TDBCombo tdbPayFreqList 
                     Height          =   345
                     Left            =   1215
                     TabIndex        =   0
                     Tag             =   "Municipal"
                     Top             =   210
                     Width           =   1995
                     _ExtentX        =   3519
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
                     MaxComboItems   =   4
                     AddItemSeparator=   ";"
                     _PropDict       =   $"frmMDPayrollPeriod.frx":4AD6
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
                  Begin TrueOleDBList80.TDBCombo tdbTmp 
                     Bindings        =   "frmMDPayrollPeriod.frx":4B80
                     DataMember      =   "tdbJob"
                     Height          =   300
                     Left            =   3480
                     TabIndex        =   48
                     Top             =   225
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
                     _PropDict       =   $"frmMDPayrollPeriod.frx":4B91
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
                  Begin VB.Label Label11 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Periods"
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
                     Left            =   120
                     TabIndex        =   38
                     Top             =   630
                     Width           =   810
                  End
                  Begin VB.Label Label10 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Filter"
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
                     Left            =   120
                     TabIndex        =   37
                     Top             =   270
                     Width           =   810
                  End
               End
               Begin VB.Frame fraSSS 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "SSS Contribution"
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
                  Height          =   1140
                  Left            =   405
                  TabIndex        =   34
                  Top             =   3900
                  Width           =   3000
                  Begin VB.CheckBox chkSSSDaily 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Daily Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   10
                     Top             =   315
                     Width           =   2520
                  End
                  Begin VB.CheckBox chkSSSMonthly 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Monthly Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   11
                     Top             =   615
                     Width           =   2520
                  End
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "PhilHealth Contribution"
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
                  Height          =   1140
                  Index           =   1
                  Left            =   390
                  TabIndex        =   33
                  Top             =   5070
                  Width           =   3000
                  Begin VB.CheckBox chkPHMonthly 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Monthly Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   15
                     Top             =   615
                     Width           =   2520
                  End
                  Begin VB.CheckBox chkPHDaily 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Daily Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   14
                     Top             =   315
                     Width           =   2520
                  End
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Withholding Tax"
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
                  Height          =   1140
                  Left            =   3435
                  TabIndex        =   32
                  Top             =   3900
                  Width           =   3000
                  Begin VB.CheckBox chkTaxDaily 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Daily Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   12
                     Top             =   315
                     Width           =   2520
                  End
                  Begin VB.CheckBox chkTaxMonthly 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Monthly Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   13
                     Top             =   615
                     Width           =   2520
                  End
               End
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "HDMF Contribution"
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
                  Height          =   1140
                  Left            =   3420
                  TabIndex        =   31
                  Top             =   5070
                  Width           =   3000
                  Begin VB.CheckBox chkHDMFMonthly 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Monthly Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   17
                     Top             =   615
                     Width           =   2520
                  End
                  Begin VB.CheckBox chkHDMFDaily 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Daily Rate Employees"
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
                     Height          =   315
                     Left            =   255
                     TabIndex        =   16
                     Top             =   315
                     Width           =   2520
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Loan Deductions"
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
                  Height          =   6390
                  Left            =   6870
                  TabIndex        =   29
                  Top             =   150
                  Width           =   5895
                  Begin TrueOleDBGrid80.TDBGrid tdgLeaveLimit 
                     Height          =   6090
                     Left            =   60
                     TabIndex        =   30
                     Top             =   225
                     Width           =   5760
                     _ExtentX        =   10160
                     _ExtentY        =   10742
                     _LayoutType     =   4
                     _RowHeight      =   16
                     _WasPersistedAsPixels=   0
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).Caption=   "Loan Types"
                     Columns(0).DataField=   "loantypesname"
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns(1)._VlistStyle=   4
                     Columns(1)._MaxComboItems=   5
                     Columns(1).DataField=   "allow"
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
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=8493"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8414"
                     Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                     Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
                     Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(7)=   "Column(1).Width=503"
                     Splits(0)._ColumnProps(8)=   "Column(1).DividerStyle=0"
                     Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                     Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=450"
                     Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                     Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
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
                     _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13,.alignment=2,.locked=0"
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
               End
               Begin VB.Frame Frame6 
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   2460
                  Left            =   60
                  TabIndex        =   39
                  Top             =   1110
                  Width           =   6780
                  Begin TrueOleDBList80.TDBCombo tdbMonth 
                     Height          =   345
                     Left            =   1215
                     TabIndex        =   6
                     Tag             =   "Municipal"
                     Top             =   1335
                     Width           =   1995
                     _ExtentX        =   3519
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
                     _PropDict       =   $"frmMDPayrollPeriod.frx":4C3B
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
                  Begin TrueOleDBList80.TDBCombo tdbClassification 
                     Height          =   345
                     Left            =   4680
                     TabIndex        =   7
                     Tag             =   "Municipal"
                     Top             =   1335
                     Width           =   2010
                     _ExtentX        =   3545
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
                     _PropDict       =   $"frmMDPayrollPeriod.frx":4CE5
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
                  Begin TrueOleDBList80.TDBCombo tdbPayFrequency 
                     Height          =   345
                     Left            =   4680
                     TabIndex        =   5
                     Tag             =   "Municipal"
                     Top             =   960
                     Width           =   2010
                     _ExtentX        =   3545
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
                     _PropDict       =   $"frmMDPayrollPeriod.frx":4D8F
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
                  Begin TDBText6Ctl.TDBText txtPayPeriod 
                     Height          =   300
                     Left            =   1215
                     TabIndex        =   2
                     Top             =   300
                     Width           =   1995
                     _Version        =   65536
                     _ExtentX        =   3519
                     _ExtentY        =   529
                     Caption         =   "frmMDPayrollPeriod.frx":4E39
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":4EA5
                     Key             =   "frmMDPayrollPeriod.frx":4EC3
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
                  Begin TDBDate6Ctl.TDBDate txtTo 
                     Height          =   300
                     Left            =   1215
                     TabIndex        =   9
                     Top             =   2040
                     Width           =   1995
                     _Version        =   65536
                     _ExtentX        =   3519
                     _ExtentY        =   529
                     Calendar        =   "frmMDPayrollPeriod.frx":4F07
                     Caption         =   "frmMDPayrollPeriod.frx":500D
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":5073
                     Keys            =   "frmMDPayrollPeriod.frx":5091
                     Spin            =   "frmMDPayrollPeriod.frx":50EF
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
                     Text            =   "01/16/2008"
                     ValidateMode    =   0
                     ValueVT         =   7
                     Value           =   39463
                     CenturyMode     =   0
                  End
                  Begin TDBText6Ctl.TDBText txtDescription 
                     Height          =   300
                     Left            =   1215
                     TabIndex        =   3
                     Top             =   630
                     Width           =   3450
                     _Version        =   65536
                     _ExtentX        =   6085
                     _ExtentY        =   529
                     Caption         =   "frmMDPayrollPeriod.frx":5117
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":5183
                     Key             =   "frmMDPayrollPeriod.frx":51A1
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
                  Begin TDBDate6Ctl.TDBDate txtFrom 
                     Height          =   300
                     Left            =   1215
                     TabIndex        =   8
                     Top             =   1710
                     Width           =   1995
                     _Version        =   65536
                     _ExtentX        =   3519
                     _ExtentY        =   529
                     Calendar        =   "frmMDPayrollPeriod.frx":51E5
                     Caption         =   "frmMDPayrollPeriod.frx":52EB
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":5351
                     Keys            =   "frmMDPayrollPeriod.frx":536F
                     Spin            =   "frmMDPayrollPeriod.frx":53CD
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
                     Text            =   "01/16/2008"
                     ValidateMode    =   0
                     ValueVT         =   7
                     Value           =   39463
                     CenturyMode     =   0
                  End
                  Begin TDBDate6Ctl.TDBDate txtPayyear 
                     Height          =   345
                     Left            =   1215
                     TabIndex        =   4
                     Top             =   960
                     Width           =   1995
                     _Version        =   65536
                     _ExtentX        =   3519
                     _ExtentY        =   609
                     Calendar        =   "frmMDPayrollPeriod.frx":53F5
                     Caption         =   "frmMDPayrollPeriod.frx":54E1
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDPayrollPeriod.frx":5547
                     Keys            =   "frmMDPayrollPeriod.frx":5565
                     Spin            =   "frmMDPayrollPeriod.frx":55C3
                     AlignHorizontal =   0
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     CursorPosition  =   0
                     DataProperty    =   0
                     DisplayFormat   =   "yyyy"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     FirstMonth      =   4
                     ForeColor       =   -2147483640
                     Format          =   "yyyy"
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
                     Text            =   "2008"
                     ValidateMode    =   0
                     ValueVT         =   2118189063
                     Value           =   39463
                     CenturyMode     =   0
                  End
                  Begin VB.Label label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Payroll Year"
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
                     Index           =   0
                     Left            =   -420
                     TabIndex        =   47
                     Top             =   1035
                     Width           =   1560
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Month"
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
                     Index           =   0
                     Left            =   -450
                     TabIndex        =   46
                     Top             =   1365
                     Width           =   1560
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Week"
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
                     Left            =   3360
                     TabIndex        =   45
                     Top             =   1365
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pay Frequency"
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
                     Left            =   3300
                     TabIndex        =   44
                     Top             =   1035
                     Width           =   1305
                  End
                  Begin VB.Label Label9 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Description"
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
                     Left            =   -465
                     TabIndex        =   43
                     Top             =   660
                     Width           =   1560
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Period: From"
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
                     Index           =   1
                     Left            =   -750
                     TabIndex        =   42
                     Top             =   1755
                     Width           =   1845
                  End
                  Begin VB.Label Label3 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Period Code"
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
                     Left            =   -465
                     TabIndex        =   41
                     Top             =   330
                     Width           =   1560
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "To"
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
                     Left            =   -135
                     TabIndex        =   40
                     Top             =   2070
                     Width           =   1230
                  End
               End
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgBrowsePic 
      Left            =   105
      Top             =   7530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBrowseSig 
      Left            =   555
      Top             =   7530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMDPayrollPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mSave                         As Boolean
Dim rsPayrollPeriod               As ADODB.Recordset
Dim rsLoanDed                     As ADODB.Recordset
Dim mTxt                          As TDBText

Private Sub ClearText()

    mSave = True
    txtPayYear.Text = ""
    tdbMonth.BoundText = ""
    tdbPayFrequency.BoundText = ""
    tdbClassification.BoundText = ""
    txtPayPeriod.Text = ""
    txtDescription.Text = ""
    txtFrom.Text = ""
    txtTo.Text = ""
    chkSSSDaily.Value = 0
    chkSSSMonthly.Value = 0
    chkPHDaily.Value = 0
    chkPHMonthly.Value = 0
    chkTaxDaily.Value = 0
    chkTaxMonthly.Value = 0
    chkHDMFDaily.Value = 0
    chkHDMFMonthly.Value = 0
    Create_TmpLoanDed 0
    
End Sub

Private Sub cmdmenu_Click(Index As Integer)
  
  Select Case Index
    Case 0:
              If MsgBox("Do you want to create a new payroll period?", vbQuestion + vbYesNo) = vbYes Then
                ClearText
              End If
    Case 1: Save_Update
    Case 2: 'Delete_Employee
    Case 3: Unload Me
  End Select
  
End Sub

Private Sub cmdSearch_Click()
  
  If Trim(tdbPayFreqList.Text) = "" Or IsNull(tdbPayFreqList.SelectedItem) Or tdbPayFreqList.ApproxCount <= 0 Then
    MsgBox "Please select a filter.", vbExclamation + vbOKOnly
    tdbPayFreqList.SetFocus
    Exit Sub
  End If
  
  bind_tdb ConMain, tdbTmp, "select percode,description from payrollperiod " & _
              "where payfreqcode = " & tdbPayFreqList.BoundText & " order by percode desc", "description", "percode"
              
  Set mTxt = txtSearch
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
'  SendKeys "{F4}"
  
End Sub

Private Sub Save_Update()


    If Trim(txtDescription.Text) = "" Then
      MsgBox "Please provide a description.", vbExclamation + vbOKOnly
      txtDescription.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtPayYear.Value) Then
      MsgBox "Invalid year format.", vbExclamation + vbOKOnly
      txtPayYear.SetFocus
      Exit Sub
    End If
     
    If Trim(tdbMonth.Text) = "" Or IsNull(tdbMonth.SelectedItem) Or tdbMonth.ApproxCount = 0 Then
      MsgBox "Please select a month.", vbExclamation + vbOKOnly
      tdbMonth.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbPayFrequency.Text) = "" Or IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
      MsgBox "Please select a pay frequency.", vbExclamation + vbOKOnly
      tdbPayFrequency.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbClassification.Text) = "" Or IsNull(tdbClassification.SelectedItem) Or tdbClassification.ApproxCount = 0 Then
      MsgBox "Please select a classification.", vbExclamation + vbOKOnly
      tdbClassification.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtFrom.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtFrom.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtTo.Text) Then
      MsgBox "Invalid date format.", vbExclamation + vbOKOnly
      txtTo.SetFocus
      Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    If mSave Then
      
      txtPayPeriod.Text = LastCode("PayrollPeriod")
      ConMain.Execute "insert into payrollperiod(payyear,paymonth,percode,description,payfreqcode," & _
            "classification,wrkdatefrom,wrkdateto,sssdaily,sssmonthly," & _
            "phdaily,phmonthly,taxdaily,taxmonthly,hdmfdaily," & _
            "hdmfmonthly,lastotcode) values " & _
            "('" & txtPayYear.Text & "','" & tdbMonth.Text & "'," & txtPayPeriod.Text & ",'" & txtDescription.Text & "','" & tdbPayFrequency.BoundText & "', " & _
            "'" & tdbClassification.Text & "','" & Format(txtFrom.Text, "YYYY-MM-DD") & "','" & Format(txtTo.Text, "YYYY-MM-DD") & "','" & IIf(chkSSSDaily.Value = 0, "N", "Y") & "','" & IIf(chkSSSMonthly.Value = 0, "N", "Y") & "'," & _
            "'" & IIf(chkPHDaily.Value = 0, "N", "Y") & "','" & IIf(chkPHMonthly.Value = 0, "N", "Y") & "','" & IIf(chkTaxDaily.Value = 0, "N", "Y") & "','" & IIf(chkTaxMonthly.Value = 0, "N", "Y") & "','" & IIf(chkHDMFDaily.Value = 0, "N", "Y") & "','" & _
            IIf(chkHDMFMonthly.Value = 0, "N", "Y") & "',1) "
                      
    Else
      
      ConMain.Execute "update payrollperiod set payyear = '" & txtPayYear.Text & "', paymonth = '" & tdbMonth.Text & "', payfreqcode = '" & tdbPayFrequency.BoundText & "', description = '" & txtDescription.Text & "'," & _
            "classification = '" & tdbClassification.Text & "',  wrkdatefrom = '" & Format(txtFrom.Text, "YYYY-MM-DD") & "', wrkdateto = '" & Format(txtTo.Text, "YYYY-MM-DD") & "', " & _
            "sssdaily = '" & IIf(chkSSSDaily.Value = 0, "N", "Y") & "', sssmonthly = '" & IIf(chkSSSMonthly.Value = 0, "N", "Y") & "', " & _
            "phdaily = '" & IIf(chkPHDaily.Value = 0, "N", "Y") & "', phmonthly = '" & IIf(chkPHMonthly.Value = 0, "N", "Y") & "', " & _
            "taxdaily = '" & IIf(chkTaxDaily.Value = 0, "N", "Y") & "', taxmonthly = '" & IIf(chkTaxMonthly.Value = 0, "N", "Y") & "', " & _
            "hdmfdaily = '" & IIf(chkHDMFDaily.Value = 0, "N", "Y") & "', hdmfmonthly = '" & IIf(chkHDMFMonthly.Value = 0, "N", "Y") & "' where percode = " & txtPayPeriod.Text & ""
                       
    End If
    
    ConMain.Execute "delete from payrollperiodloandedallow where percode = " & txtPayPeriod.Text & ""
    
    With rsLoanDed
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !allow <> 0 Then
                    ConMain.Execute "insert into payrollperiodloandedallow (percode,loantypescode,allow) values (" & _
                             txtPayPeriod.Text & "," & !loantypescode & "," & IIf(!allow <> 0, 1, 0) & ")"
                End If
                .MoveNext
            Loop
        End If
    End With
            
      
    ConMain.CommitTrans
        
    If mSave Then
      mSave = False
      MsgBox "Payroll period has been successfully saved.", vbInformation + vbOKOnly
    Else
      MsgBox "Payroll period has been successfully updated.", vbInformation + vbOKOnly
    End If

End Sub
Private Sub Form_Activate()

    Focus_MDIButton Me

End Sub

Private Sub Form_Load()

    Dim rsTmpPayFreq  As ADODB.Recordset
    Dim rsTmp         As ADODB.Recordset
    Dim i             As Integer
    
    Add_MDIButton Me.Name, Me.Tag
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 12
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "January"
          Case 2: .Fields("description") = "February"
          Case 3: .Fields("description") = "March"
          Case 4: .Fields("description") = "April"
          Case 5: .Fields("description") = "May"
          Case 6: .Fields("description") = "June"
          Case 7: .Fields("description") = "July"
          Case 8: .Fields("description") = "August"
          Case 9: .Fields("description") = "September"
          Case 10: .Fields("description") = "October"
          Case 11: .Fields("description") = "November"
          Case 12: .Fields("description") = "December"
        End Select
        .Update
      Next
    End With
    
    With tdbMonth
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 5
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "First"
          Case 2: .Fields("description") = "Second"
          Case 3: .Fields("description") = "Third"
          Case 4: .Fields("description") = "Fourth"
          Case 5: .Fields("description") = "Fifth"
        End Select
        .Update
      Next
    End With
    
    With tdbClassification
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    NetOpen rsTmpPayFreq, "select * from payfrequency order by payfreqname"
    With rsTmpPayFreq
      CreateTmpDB rsTmp
      If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
          rsTmp.AddNew
          rsTmp.Fields("code") = !payfreqcode
          rsTmp.Fields("Description") = !payfreqname
          rsTmp.Update
          .MoveNext
        Loop
      End If
    End With

    With tdbPayFreqList
      .BoundColumn = "code"
      .ListField = "description"
      .Columns(0).DataField = "code"
      .Columns(1).DataField = "description"
      .RowSource = rsTmp
    End With
      
    Set rsTmp = Nothing
    
    bind_tdb ConMain, tdbPayFrequency, "select payfreqcode,payfreqname from payfrequency order by payfreqname", "payfreqname", "payfreqcode"

    ClearText
    
End Sub

Private Sub Form_Resize()
  
  On Error Resume Next
  
  With fraHolder
    .Top = (Me.ScaleHeight / 2) - (.Height / 2) + ((pic1.Top + pic1.Height) / 2)
    .Left = (Me.ScaleWidth / 2) - (.Width / 2)
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub tdbClassification_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbClassification, tdbClassification.RowSource, tdbClassification.Text
  End If
End Sub

Private Sub tdbClassification_GotFocus()
  With tdbClassification
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbMonth_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbMonth, tdbMonth.RowSource, tdbMonth.Text
  End If
End Sub

Private Sub tdbMonth_Gotfocus()
  With tdbMonth
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbPayFreqList_GotFocus()
  With tdbPayFreqList
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbPayFreqList_ItemChange()
  txtSearch.Tag = ""
  txtSearch.Text = ""
End Sub

Private Sub tdbPayFreqList_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbPayFreqList, tdbPayFreqList.RowSource, tdbPayFreqList.Text
  End If
End Sub

Private Sub tdbPayFrequency_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbPayFrequency, tdbPayFrequency.RowSource, tdbPayFrequency.Text
  End If
End Sub

Private Sub tdbPayFrequency_GotFocus()
  With tdbPayFrequency
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
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
  
  If Trim(tdbTmp.Text) <> "" And Not IsNull(tdbTmp.SelectedItem) And tdbTmp.ApproxCount > 0 Then
    
    NetOpen rsPayrollPeriod, "select x1.*,x2.payfreqname from payrollperiod x1 " & _
                               "left outer join payfrequency x2 on x1.payfreqcode = x2.payfreqcode where x1.percode = " & tdbTmp.BoundText & ""
                               
    With rsPayrollPeriod
      If .RecordCount > 0 Then
          txtPayYear.Text = !payyear
          tdbMonth.Text = !paymonth
          tdbPayFrequency.BoundText = !payfreqcode
          tdbClassification.Text = !classification
          txtPayPeriod.Text = !percode
          txtDescription.Text = !Description
          txtFrom.Text = Format(!wrkdatefrom, "MM/DD/YYYY")
          txtTo.Text = Format(!wrkdateto, "MM/DD/YYYY")
          chkSSSDaily.Value = IIf(!sssdaily = "N", 0, 1)
          chkSSSMonthly.Value = IIf(!sssmonthly = "N", 0, 1)
          chkPHDaily.Value = IIf(!phdaily = "N", 0, 1)
          chkPHMonthly.Value = IIf(!phmonthly = "N", 0, 1)
          chkTaxDaily.Value = IIf(!taxdaily = "N", 0, 1)
          chkTaxMonthly.Value = IIf(!taxmonthly = "N", 0, 1)
          chkHDMFDaily.Value = IIf(!hdmfdaily = "N", 0, 1)
          chkHDMFMonthly.Value = IIf(!hdmfmonthly = "N", 0, 1)
          Create_TmpLoanDed !percode
          mSave = False
      Else
        ClearText
      End If
    End With
  Else
    ClearText
  End If
  
End Sub

Private Sub txtDescription_GotFocus()
  With txtDescription
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDescription_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtFrom_GotFocus()
  With txtFrom
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtPayPeriod_GotFocus()
  With txtPayPeriod
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPayPeriod_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtPayYear_GotFocus()
  With txtPayYear
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtPayYear_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSearch_GotFocus()
  With txtSearch
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then cmdSearch_Click
End Sub

Public Sub Create_TmpLoanDed(mPerCode As Integer)

    Dim rsLoanDedTmp        As ADODB.Recordset

    Set rsLoanDed = Nothing
    Set rsLoanDed = New ADODB.Recordset
    
    With rsLoanDed
      .Fields.Append "loantypescode", adVarChar, 7
      .Fields.Append "loantypesname", adVarChar, 50
      .Fields.Append "allow", adInteger
      .Open
    End With
    
    Set tdgLeaveLimit.DataSource = rsLoanDed
    
    NetOpen rsLoanDedTmp, "select x1.loantypescode,x1.loantypesname," & _
                        "(select allow from payrollperiodloandedallow where loantypescode = x1.loantypescode and percode = " & mPerCode & " group by loantypescode) allow " & _
                        "from loantypes x1 order by x1.loantypesname "

    With rsLoanDedTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                rsLoanDed.AddNew
                rsLoanDed.Fields("loantypescode") = !loantypescode
                rsLoanDed.Fields("loantypesname") = !loantypesname
                rsLoanDed.Fields("allow") = IIf(IsNull(!allow), 0, !allow)
                rsLoanDed.Update
                .MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub txtTo_GotFocus()
  With txtTo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub
