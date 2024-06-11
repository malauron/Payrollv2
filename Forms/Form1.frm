VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Tab tabShift 
      Height          =   5775
      Left            =   75
      TabIndex        =   0
      Top             =   1125
      Width           =   8250
      _cx             =   14552
      _cy             =   10186
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
      BackColor       =   -2147483626
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483626
      BackTabColor    =   -2147483626
      TabOutlineColor =   12632256
      FrontTabForeColor=   0
      Caption         =   ""
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
         Height          =   5460
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   8220
         _cx             =   14499
         _cy             =   9631
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
         BackColor       =   -2147483626
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
         Begin TDBText6Ctl.TDBText TDBText3 
            Height          =   300
            Left            =   1740
            TabIndex        =   2
            Top             =   435
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   529
            Caption         =   "Form1.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":006C
            Key             =   "Form1.frx":008A
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
         Begin TDBTime6Ctl.TDBTime txtT1out 
            Height          =   300
            Left            =   3960
            TabIndex        =   3
            Top             =   1380
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":00CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":013A
            Spin            =   "Form1.frx":018A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBNumber6Ctl.TDBNumber txtT1Hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   4
            Top             =   1380
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":01B2
            Caption         =   "Form1.frx":01D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":023E
            Keys            =   "Form1.frx":025C
            Spin            =   "Form1.frx":02A6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtT2hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   5
            Top             =   1710
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":02CE
            Caption         =   "Form1.frx":02EE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":035A
            Keys            =   "Form1.frx":0378
            Spin            =   "Form1.frx":03C2
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtT3Hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   6
            Top             =   2040
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":03EA
            Caption         =   "Form1.frx":040A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":0476
            Keys            =   "Form1.frx":0494
            Spin            =   "Form1.frx":04DE
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtT4Hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   7
            Top             =   2370
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":0506
            Caption         =   "Form1.frx":0526
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":0592
            Keys            =   "Form1.frx":05B0
            Spin            =   "Form1.frx":05FA
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtBrk1hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   8
            Top             =   2700
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":0622
            Caption         =   "Form1.frx":0642
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":06AE
            Keys            =   "Form1.frx":06CC
            Spin            =   "Form1.frx":0716
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtBrk2hrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   9
            Top             =   3030
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":073E
            Caption         =   "Form1.frx":075E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":07CA
            Keys            =   "Form1.frx":07E8
            Spin            =   "Form1.frx":0832
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtOtHrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   10
            Top             =   3360
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":085A
            Caption         =   "Form1.frx":087A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":08E6
            Keys            =   "Form1.frx":0904
            Spin            =   "Form1.frx":094E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtNpremhrs 
            Height          =   300
            Left            =   6150
            TabIndex        =   11
            Top             =   3690
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   529
            Calculator      =   "Form1.frx":0976
            Caption         =   "Form1.frx":0996
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":0A02
            Keys            =   "Form1.frx":0A20
            Spin            =   "Form1.frx":0A6A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "#,###,###,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "#,###,###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   -999999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBTime6Ctl.TDBTime txtT1In 
            Height          =   300
            Left            =   1770
            TabIndex        =   12
            Top             =   1380
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0A92
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0AFE
            Spin            =   "Form1.frx":0B4E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT2out 
            Height          =   300
            Left            =   3960
            TabIndex        =   13
            Top             =   1710
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0B76
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0BE2
            Spin            =   "Form1.frx":0C32
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT2in 
            Height          =   300
            Left            =   1770
            TabIndex        =   14
            Top             =   1710
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0C5A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0CC6
            Spin            =   "Form1.frx":0D16
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT3out 
            Height          =   300
            Left            =   3960
            TabIndex        =   15
            Top             =   2040
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0D3E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0DAA
            Spin            =   "Form1.frx":0DFA
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT3in 
            Height          =   300
            Left            =   1770
            TabIndex        =   16
            Top             =   2040
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0E22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0E8E
            Spin            =   "Form1.frx":0EDE
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT4out 
            Height          =   300
            Left            =   3960
            TabIndex        =   17
            Top             =   2370
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0F06
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":0F72
            Spin            =   "Form1.frx":0FC2
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtT4In 
            Height          =   300
            Left            =   1770
            TabIndex        =   18
            Top             =   2370
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":0FEA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":1056
            Spin            =   "Form1.frx":10A6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtBrk1out 
            Height          =   300
            Left            =   3960
            TabIndex        =   19
            Top             =   2700
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":10CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":113A
            Spin            =   "Form1.frx":118A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtBrk1in 
            Height          =   300
            Left            =   1770
            TabIndex        =   20
            Top             =   2700
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":11B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":121E
            Spin            =   "Form1.frx":126E
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtBrk2out 
            Height          =   300
            Left            =   3960
            TabIndex        =   21
            Top             =   3030
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":1302
            Spin            =   "Form1.frx":1352
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtBrk2in 
            Height          =   300
            Left            =   1770
            TabIndex        =   22
            Top             =   3030
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":137A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":13E6
            Spin            =   "Form1.frx":1436
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtOvertimeout 
            Height          =   300
            Left            =   3960
            TabIndex        =   23
            Top             =   3360
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":145E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":14CA
            Spin            =   "Form1.frx":151A
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtOvertimeIn 
            Height          =   300
            Left            =   1770
            TabIndex        =   24
            Top             =   3360
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":1542
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":15AE
            Spin            =   "Form1.frx":15FE
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtNPremOut 
            Height          =   300
            Left            =   3960
            TabIndex        =   25
            Top             =   3690
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":1626
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":1692
            Spin            =   "Form1.frx":16E2
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin TDBTime6Ctl.TDBTime txtNPremIn 
            Height          =   300
            Left            =   1770
            TabIndex        =   26
            Top             =   3690
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "Form1.frx":170A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "Form1.frx":1776
            Spin            =   "Form1.frx":17C6
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   2
            BackColor       =   -2147483643
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "hh:nn am/pm"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "hh:nn am/pm"
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
            Text            =   "10:45 pm"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.948009259259259
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Shift Code"
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
            Left            =   -45
            TabIndex        =   38
            Top             =   375
            Width           =   1560
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "HOURS"
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
            Left            =   6480
            TabIndex        =   37
            Top             =   990
            Width           =   915
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Night Premium"
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
            Left            =   150
            TabIndex        =   36
            Top             =   3720
            Width           =   1470
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime"
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
            Left            =   375
            TabIndex        =   35
            Top             =   3390
            Width           =   1230
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "2nd Break"
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
            Left            =   390
            TabIndex        =   34
            Top             =   3075
            Width           =   1230
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "1st Break"
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
            Left            =   390
            TabIndex        =   33
            Top             =   2730
            Width           =   1230
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TIME OUT"
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
            Left            =   4470
            TabIndex        =   32
            Top             =   990
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TIME IN"
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
            Left            =   2250
            TabIndex        =   31
            Top             =   990
            Width           =   750
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "4th TITO"
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
            Left            =   375
            TabIndex        =   30
            Top             =   2415
            Width           =   1230
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "2nd TITO"
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
            Left            =   45
            TabIndex        =   29
            Top             =   1770
            Width           =   1560
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "3rd TITO"
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
            TabIndex        =   28
            Top             =   2100
            Width           =   1845
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "1st TITO"
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
            Left            =   765
            TabIndex        =   27
            Top             =   1440
            Width           =   855
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   5460
         Left            =   8865
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   300
         Width           =   8220
         _cx             =   14499
         _cy             =   9631
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
         BackColor       =   -2147483626
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
         Begin VB.Frame fraSearch 
            BackColor       =   &H80000016&
            Height          =   720
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   5895
            Begin TDBText6Ctl.TDBText txtSearch 
               Height          =   300
               Left            =   1395
               TabIndex        =   41
               Top             =   255
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
               _ExtentY        =   529
               Caption         =   "Form1.frx":17EE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form1.frx":185A
               Key             =   "Form1.frx":1878
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
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "SEARCH"
               Height          =   255
               Left            =   375
               TabIndex        =   42
               Top             =   315
               Width           =   915
            End
         End
         Begin TrueOleDBGrid80.TDBGrid tdgShift 
            Height          =   4155
            Left            =   60
            TabIndex        =   43
            Top             =   1260
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   7329
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Code"
            Columns(0).DataField=   "shiftcode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "T1in"
            Columns(1).DataField=   "t1in"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "T1out"
            Columns(2).DataField=   "t1out"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T2in"
            Columns(3).DataField=   "T2in"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "T2out"
            Columns(4).DataField=   "t2out"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T3in"
            Columns(5).DataField=   "T3out"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T3out"
            Columns(6).DataField=   "t3out"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "T4in"
            Columns(7).DataField=   "t4out"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "T4out"
            Columns(8).DataField=   "t4out"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "1st Break From"
            Columns(9).DataField=   "brk1in"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "1st Break To"
            Columns(10).DataField=   "brk1out"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "2nd Break From"
            Columns(11).DataField=   "brk2in"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "2nd Break To"
            Columns(12).DataField=   "brk2out"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Overtime in"
            Columns(13).DataField=   "OvertimeIn"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Overtime out"
            Columns(14).DataField=   "overtimeout"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Night Premium from"
            Columns(15).DataField=   "npremfrom"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Night Premium to"
            Columns(16).DataField=   ""
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   17
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=17"
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
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
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(51)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(54)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(56)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(59)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(61)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(64)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(66)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(69)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(71)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(74)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(76)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(79)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(81)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(84)=   "Column(16)._EditAlways=0"
            Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
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
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(53)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(54)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(55)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(56)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(57)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(58)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(59)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(60)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(61)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(62)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(63)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(64)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(65)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(66)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(67)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(68)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(69)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(70)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(71)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(72)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(73)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(74)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(75)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(76)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(77)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(78)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(79)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(80)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(81)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
            _StyleDefs(82)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(83)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(84)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(85)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
            _StyleDefs(86)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
            _StyleDefs(87)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
            _StyleDefs(88)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
            _StyleDefs(89)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
            _StyleDefs(90)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
            _StyleDefs(91)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
            _StyleDefs(92)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
            _StyleDefs(93)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
            _StyleDefs(94)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
            _StyleDefs(95)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
            _StyleDefs(96)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
            _StyleDefs(97)  =   "Splits(0).Columns(16).Style:id=102,.parent=13"
            _StyleDefs(98)  =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
            _StyleDefs(99)  =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
            _StyleDefs(100) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
            _StyleDefs(101) =   "Named:id=33:Normal"
            _StyleDefs(102) =   ":id=33,.parent=0"
            _StyleDefs(103) =   "Named:id=34:Heading"
            _StyleDefs(104) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(105) =   ":id=34,.wraptext=-1"
            _StyleDefs(106) =   "Named:id=35:Footing"
            _StyleDefs(107) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(108) =   "Named:id=36:Selected"
            _StyleDefs(109) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(110) =   "Named:id=37:Caption"
            _StyleDefs(111) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(112) =   "Named:id=38:HighlightRow"
            _StyleDefs(113) =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(114) =   "Named:id=39:EvenRow"
            _StyleDefs(115) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(116) =   "Named:id=40:OddRow"
            _StyleDefs(117) =   ":id=40,.parent=33"
            _StyleDefs(118) =   "Named:id=41:RecordSelector"
            _StyleDefs(119) =   ":id=41,.parent=34"
            _StyleDefs(120) =   "Named:id=42:FilterBar"
            _StyleDefs(121) =   ":id=42,.parent=33"
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   5460
         Left            =   9165
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   300
         Width           =   8220
         _cx             =   14499
         _cy             =   9631
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
            TabIndex        =   45
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   46
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "Form1.frx":18BC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form1.frx":1928
               Key             =   "Form1.frx":1946
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
               TabIndex        =   47
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "Form1.frx":198A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form1.frx":19F6
               Key             =   "Form1.frx":1A14
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
               TabIndex        =   48
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "Form1.frx":1A58
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Form1.frx":1AC4
               Key             =   "Form1.frx":1AE2
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
               Top             =   945
               Width           =   990
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   52
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "Form1.frx":1B26
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form1.frx":1B92
            Key             =   "Form1.frx":1BB0
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
            TabIndex        =   53
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
            FormatString    =   $"Form1.frx":1BF4
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
            TabIndex        =   54
            Top             =   240
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
