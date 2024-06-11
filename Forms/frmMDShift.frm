VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDShift 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab tabShift 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   750
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
      BackColor       =   16185592
      ForeColor       =   -2147483630
      FrontTabColor   =   16185592
      BackTabColor    =   16185592
      TabOutlineColor =   12632256
      FrontTabForeColor=   0
      Caption         =   "Maintain Shift Schedule|View Shift Schedules"
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
         Begin VB.Frame FraInfo 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Height          =   4095
            Left            =   105
            TabIndex        =   14
            Top             =   90
            Width           =   8055
            Begin VB.CheckBox chkRequired 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Required"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4395
               TabIndex        =   51
               Top             =   315
               Width           =   2325
            End
            Begin TDBText6Ctl.TDBText txtShiftcode 
               Height          =   300
               Left            =   2415
               TabIndex        =   15
               Top             =   300
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0000
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":006C
               Key             =   "frmMDShift.frx":008A
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
            Begin TDBTime6Ctl.TDBTime txtT1out 
               Height          =   300
               Left            =   4275
               TabIndex        =   16
               Top             =   1245
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":00CE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":013A
               Spin            =   "frmMDShift.frx":018A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
               Left            =   6165
               TabIndex        =   17
               Top             =   1245
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   529
               Calculator      =   "frmMDShift.frx":01B2
               Caption         =   "frmMDShift.frx":01D2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":023E
               Keys            =   "frmMDShift.frx":025C
               Spin            =   "frmMDShift.frx":02A6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBNumber6Ctl.TDBNumber txtT2hrs 
               Height          =   300
               Left            =   6165
               TabIndex        =   18
               Top             =   1575
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   529
               Calculator      =   "frmMDShift.frx":02CE
               Caption         =   "frmMDShift.frx":02EE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":035A
               Keys            =   "frmMDShift.frx":0378
               Spin            =   "frmMDShift.frx":03C2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBNumber6Ctl.TDBNumber txtBrkhrs 
               Height          =   300
               Left            =   6180
               TabIndex        =   19
               Top             =   1905
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Calculator      =   "frmMDShift.frx":03EA
               Caption         =   "frmMDShift.frx":040A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0476
               Keys            =   "frmMDShift.frx":0494
               Spin            =   "frmMDShift.frx":04DE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBNumber6Ctl.TDBNumber txtNitepremHrs 
               Height          =   300
               Left            =   6165
               TabIndex        =   20
               Top             =   2985
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Calculator      =   "frmMDShift.frx":0506
               Caption         =   "frmMDShift.frx":0526
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0592
               Keys            =   "frmMDShift.frx":05B0
               Spin            =   "frmMDShift.frx":05FA
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBTime6Ctl.TDBTime txtT1In 
               Height          =   300
               Left            =   2415
               TabIndex        =   21
               Top             =   1245
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0622
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":068E
               Spin            =   "frmMDShift.frx":06DE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
               Left            =   4275
               TabIndex        =   22
               Top             =   1575
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0706
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":0772
               Spin            =   "frmMDShift.frx":07C2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
               Left            =   2415
               TabIndex        =   23
               Top             =   1575
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":07EA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":0856
               Spin            =   "frmMDShift.frx":08A6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBTime6Ctl.TDBTime txtBrkend 
               Height          =   300
               Left            =   4275
               TabIndex        =   24
               Top             =   2655
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":08CE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":093A
               Spin            =   "frmMDShift.frx":098A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBTime6Ctl.TDBTime txtBrkstart 
               Height          =   300
               Left            =   2415
               TabIndex        =   25
               Top             =   2655
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":09B2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":0A1E
               Spin            =   "frmMDShift.frx":0A6E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBTime6Ctl.TDBTime txtNitePremend 
               Height          =   300
               Left            =   4275
               TabIndex        =   38
               Top             =   2985
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0A96
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":0B02
               Spin            =   "frmMDShift.frx":0B52
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin TDBTime6Ctl.TDBTime txtNitePremstart 
               Height          =   300
               Left            =   2415
               TabIndex        =   39
               Top             =   2985
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0B7A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmMDShift.frx":0BE6
               Spin            =   "frmMDShift.frx":0C36
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "FROM"
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
               Left            =   2790
               TabIndex        =   42
               Top             =   2340
               Width           =   750
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "TO"
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
               Left            =   4380
               TabIndex        =   41
               Top             =   2340
               Width           =   915
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nite Premium hours"
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
               Left            =   75
               TabIndex        =   40
               Top             =   3045
               Width           =   2025
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Break Hours"
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
               Left            =   3975
               TabIndex        =   37
               Top             =   1965
               Width           =   2025
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
               Left            =   495
               TabIndex        =   32
               Top             =   315
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
               Left            =   6345
               TabIndex        =   31
               Top             =   855
               Width           =   915
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Allowable break hours"
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
               Left            =   75
               TabIndex        =   30
               Top             =   2715
               Width           =   2025
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
               Left            =   4650
               TabIndex        =   29
               Top             =   855
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
               Left            =   2865
               TabIndex        =   28
               Top             =   855
               Width           =   750
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
               Left            =   540
               TabIndex        =   27
               Top             =   1635
               Width           =   1560
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
               Left            =   1245
               TabIndex        =   26
               Top             =   1305
               Width           =   855
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   5460
         Left            =   8865
         TabIndex        =   2
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
         Begin VB.Frame fraSearch 
            BackColor       =   &H00F6F8F8&
            Height          =   720
            Left            =   60
            TabIndex        =   34
            Top             =   90
            Width           =   8175
            Begin TDBText6Ctl.TDBText txtSearch 
               Height          =   300
               Left            =   1395
               TabIndex        =   35
               Top             =   255
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7064
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0C5E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0CCA
               Key             =   "frmMDShift.frx":0CE8
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
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "SEARCH"
               Height          =   255
               Left            =   375
               TabIndex        =   36
               Top             =   315
               Width           =   915
            End
         End
         Begin TrueOleDBGrid80.TDBGrid tdgShift 
            Height          =   4530
            Left            =   -2925
            TabIndex        =   33
            Top             =   930
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   7990
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
            Columns(1).Caption=   "1st Time in"
            Columns(1).DataField=   "t1in"
            Columns(1).NumberFormat=   "hh:nn am/pm"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "1st Time out"
            Columns(2).DataField=   "t1out"
            Columns(2).NumberFormat=   "hh:nn am/pm"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "2nd Time in"
            Columns(3).DataField=   "t2in"
            Columns(3).NumberFormat=   "hh:nn am/pm"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "2nd Time out"
            Columns(4).DataField=   "t2out"
            Columns(4).NumberFormat=   "hh:nn am/pm"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Allowable Break from"
            Columns(5).DataField=   "brkstart"
            Columns(5).NumberFormat=   "hh:nn am/pm"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Allowable Break to"
            Columns(6).DataField=   "brkend"
            Columns(6).NumberFormat=   "hh:nn am/pm"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Night Premium In"
            Columns(7).DataField=   "nitepremstart"
            Columns(7).NumberFormat=   "hh:nn am/pm"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Night Premium Out"
            Columns(8).DataField=   "nitepremend"
            Columns(8).NumberFormat=   "hh:nn am/pm"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
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
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1799"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1720"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=1720"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1640"
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
            HeadLines       =   2
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=86,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=90,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=110,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=114,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=111,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=112,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=113,.parent=17"
            _StyleDefs(70)  =   "Named:id=33:Normal"
            _StyleDefs(71)  =   ":id=33,.parent=0"
            _StyleDefs(72)  =   "Named:id=34:Heading"
            _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   ":id=34,.wraptext=-1"
            _StyleDefs(75)  =   "Named:id=35:Footing"
            _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=36:Selected"
            _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=37:Caption"
            _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(81)  =   "Named:id=38:HighlightRow"
            _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(83)  =   "Named:id=39:EvenRow"
            _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(85)  =   "Named:id=40:OddRow"
            _StyleDefs(86)  =   ":id=40,.parent=33"
            _StyleDefs(87)  =   "Named:id=41:RecordSelector"
            _StyleDefs(88)  =   ":id=41,.parent=34"
            _StyleDefs(89)  =   "Named:id=42:FilterBar"
            _StyleDefs(90)  =   ":id=42,.parent=33"
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   5460
         Left            =   9165
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   5
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0D2C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0D98
               Key             =   "frmMDShift.frx":0DB6
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
               TabIndex        =   6
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0DFA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0E66
               Key             =   "frmMDShift.frx":0E84
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
               TabIndex        =   7
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDShift.frx":0EC8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDShift.frx":0F34
               Key             =   "frmMDShift.frx":0F52
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
               Top             =   615
               Width           =   1635
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   11
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "frmMDShift.frx":0F96
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDShift.frx":1002
            Key             =   "frmMDShift.frx":1020
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
            TabIndex        =   12
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
            FormatString    =   $"frmMDShift.frx":1064
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
            TabIndex        =   13
            Top             =   240
            Width           =   915
         End
      End
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
      TabIndex        =   50
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
Attribute VB_Name = "frmMDShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rShift          As ADODB.Recordset
Dim mSort           As String

Private Sub cmdmenu_Click(Index As Integer)
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

    Add_MDIButton Me.Name, TitleBar.Caption
    
    NetOpen rShift, "select* from shift order by shiftcode"
    If rShift.RecordCount > 0 Then
      rShift.MoveFirst
    End If
    Set tdgShift.DataSource = rShift
    mSort = "shiftcode"
    tabShift.CurrTab = 0
    cmdmenu_Click 3

End Sub

Private Sub AddSave_Button_Clicked()

    Dim mAdvCnt     As Integer

    If cmdMenu(0).Caption = "&New" Then
    
        Lock_Button "TFFTFF", cmdMenu, 5
        cmdMenu(0).Caption = "&Save"
        ClearText
        Lock_Tab "TF", tabShift, 1
        fraInfo.Enabled = True
        tabShift.CurrTab = 0
        chkRequired.Value = 1
        txtT1In.SetFocus
    
    Else
    
        If Not IsDate(txtT1In.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT1In.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT1out.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT1out.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT2in.Text) And IsDate(txtT2out.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT2in.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT2out.Text) And IsDate(txtT2in.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT2out.SetFocus
            Exit Sub
        End If
        
        If CDate(txtT1In.Text) > CDate(txtT1out.Text) Then
            mAdvCnt = 1
        End If
        
        If IsDate(txtT2in.Text) And IsDate(txtT2out.Text) Then
            If CDate(txtT1out.Text) > CDate(txtT2in.Text) Then
                mAdvCnt = mAdvCnt + 1
                If mAdvCnt > 1 Then
                    MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
                    txtT2in.SetFocus
                    Exit Sub
                End If
            End If
            
            If CDate(txtT2in.Text) > CDate(txtT2out.Text) Then
                mAdvCnt = mAdvCnt + 1
                If mAdvCnt > 1 Then
                    MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
                    txtT2out.SetFocus
                    Exit Sub
                End If
            End If
        End If
    
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        txtShiftcode.Text = LastCode("Shift")
        ConMain.Execute "insert into shift (shiftcode,t1in,t1out,t2in,t2out,brkstart,brkend, " & _
                          "NitePremstart,NitePremend,t1hrs,t2hrs,brkhrs,nitepremhrs,required) values " & _
                          "('" & txtShiftcode.Text & "','" & IIf(CDbl(txtT1Hrs.Text) <= 0, "", Format(txtT1In.Text, "hh:nn")) & "','" & IIf(CDbl(txtT1Hrs.Text) <= 0, "", Format(txtT1out.Text, "hh:nn")) & "', " & _
                       "'" & IIf(CDbl(txtT2hrs.Text) <= 0, "", Format(txtT2in.Text, "hh:nn")) & "','" & IIf(CDbl(txtT2hrs.Text) <= 0, "", Format(txtT2out.Text, "hh:nn")) & "', " & _
                       "'" & IIf(Not IsDate(txtBrkstart.Text) Or Not IsDate(txtBrkend.Text), "", Format(txtBrkstart.Text, "hh:nn")) & "','" & IIf(Not IsDate(txtBrkstart.Text) Or Not IsDate(txtBrkend.Text), "", Format(txtBrkend.Text, "hh:nn")) & "', " & _
                       "'" & IIf(CDbl(txtNitepremHrs.Text) <= 0, "", Format(txtNitePremstart.Text, "hh:nn")) & "','" & IIf(CDbl(txtNitepremHrs.Text) <= 0, "", Format(txtNitePremend.Text, "hh:nn")) & "', " & _
                       CDbl(txtT1Hrs.Text) & "," & CDbl(txtT2hrs.Text) & "," & CDbl(txtBrkhrs.Text) & "," & CDbl(txtNitepremHrs.Text) & ",'" & IIf(chkRequired.Value = 0, "N", "Y") & "')"
        ConMain.CommitTrans
        rShift.Requery
        rShift.Find "shiftcode = '" & txtShiftcode.Text & "'"
            
        cmdmenu_Click 3
        
    End If
  
End Sub

Private Sub EditUpdate_Button_Clicked()

    Dim mAdvCnt      As Integer

    If cmdMenu(1).Caption = "&Edit" Then
      
        Lock_Button "FTFTFF", cmdMenu, 5
        cmdMenu(1).Caption = "&Update"
        Lock_Tab "TF", tabShift, 1
        fraInfo.Enabled = True
        tabShift.CurrTab = 0
        txtT1In.SetFocus
    
    Else
    
        If Not IsDate(txtT1In.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT1In.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT1out.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT1out.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT2in.Text) And IsDate(txtT2out.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT2in.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(txtT2out.Text) And IsDate(txtT2in.Text) Then
            MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
            txtT2out.SetFocus
            Exit Sub
        End If
        
        If CDate(txtT1In.Text) > CDate(txtT1out.Text) Then
            mAdvCnt = 1
        End If
        
        If IsDate(txtT2in.Text) And IsDate(txtT2out.Text) Then
            If CDate(txtT1out.Text) > CDate(txtT2in.Text) Then
                mAdvCnt = mAdvCnt + 1
                If mAdvCnt > 1 Then
                    MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
                    txtT2in.SetFocus
                    Exit Sub
                End If
            End If
            
            If CDate(txtT2in.Text) > CDate(txtT2out.Text) Then
                mAdvCnt = mAdvCnt + 1
                If mAdvCnt > 1 Then
                    MsgBox "Please enter a valid time.", vbExclamation + vbOKOnly
                    txtT2out.SetFocus
                    Exit Sub
                End If
            End If
        End If
    
        ConMain.Execute "update shift set t1in = '" & IIf(CDbl(txtT1Hrs.Text) <= 0, "", Format(txtT1In.Text, "hh:nn")) & "', t1out = '" & IIf(CDbl(txtT1Hrs.Text) <= 0, "", Format(txtT1out.Text, "hh:nn")) & "', " & _
                       " t2in = '" & IIf(CDbl(txtT2hrs.Text) <= 0, "", Format(txtT2in.Text, "hh:nn")) & "', t2out = '" & IIf(CDbl(txtT2hrs.Text) <= 0, "", Format(txtT2out.Text, "hh:nn")) & "', " & _
                       " brkstart = '" & IIf(Not IsDate(txtBrkstart.Text) Or Not IsDate(txtBrkend.Text), "", Format(txtBrkstart.Text, "hh:nn")) & "', brkend = '" & IIf(Not IsDate(txtBrkstart.Text) Or Not IsDate(txtBrkend.Text), "", Format(txtBrkend.Text, "hh:nn")) & "', " & _
                       " NitePremstart = '" & IIf(CDbl(txtNitepremHrs.Text) <= 0, "", Format(txtNitePremstart.Text, "hh:nn")) & "', NitePremend = '" & IIf(CDbl(txtNitepremHrs.Text) <= 0, "", Format(txtNitePremend.Text, "hh:nn")) & "', t1hrs = " & _
                       CDbl(txtT1Hrs.Text) & ", t2hrs = " & CDbl(txtT2hrs.Text) & ", brkhrs = " & CDbl(txtBrkhrs.Text) & ", nitepremhrs = " & CDbl(txtNitepremHrs.Text) & ", required = '" & IIf(chkRequired.Value = 0, "N", "Y") & "' where shiftcode = '" & txtShiftcode.Text & "'"
                           
        rShift.Requery
        rShift.Find "shiftcode = '" & txtShiftcode.Text & "'"
            
        cmdmenu_Click 3
      
    End If
  
End Sub

Private Sub Cancel_Clicked()

  If rShift.RecordCount > 0 Then
    Lock_Button "TTTFTT", cmdMenu, 5
  Else
    Lock_Button "TFFFTT", cmdMenu, 5
  End If

  cmdMenu(0).Caption = "&New"
  cmdMenu(1).Caption = "&Edit"
  
  Lock_Tab "TT", tabShift, 1
  fraInfo.Enabled = False
  tdgShift_RowColChange 0, 0
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
      .Top = TitleBar.Top + TitleBar.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With tabShift
      .Top = frabutton.Top + frabutton.Height
      .Left = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight - .Top
    End With
    
    With fraSearch
      .Top = 0
      .Left = 150
      .Width = Me.ScaleWidth - 300
    End With
    
    With tdgShift
      .Top = fraSearch.Height
      .Left = 150
      .Width = tabShift.Width - 300
      .Height = tabShift.Height - (.Top + 400)
    End With

End Sub

Private Sub tdgShift_HeadClick(ByVal ColIndex As Integer)
  
  If ColIndex = 0 Then
    If rShift.RecordCount > 0 Then
      mSort = tdgShift.Columns(ColIndex).DataField
      rShift.Sort = mSort
    End If
  End If
  
End Sub

Private Sub tdgShift_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

  With rShift
    If .RecordCount > 0 Then
      txtShiftcode.Text = !shiftcode
      txtT1In.Text = IIf(!t1in = "", "", Format(!t1in, "hh:nn am/pm"))
      txtT2in.Text = IIf(!t2in = "", "", Format(!t2in, "hh:nn am/pm"))
      txtT1out.Text = IIf(!t1out = "", "", Format(!t1out, "hh:nn am/pm"))
      txtT2out.Text = IIf(!t2out = "", "", Format(!t2out, "hh:nn am/pm"))
      txtBrkstart.Text = IIf(!brkstart = "", "", Format(!brkstart, "hh:nn am/pm"))
      txtBrkend.Text = IIf(!brkend = "", "", Format(!brkend, "hh:nn am/pm"))
      txtNitePremstart.Text = IIf(!nitepremstart = "", "", Format(!nitepremstart, "hh:nn am/pm"))
      txtNitePremend.Text = IIf(!nitepremend = "", "", Format(!nitepremend, "hh:nn am/pm"))
      txtT1Hrs.Text = Format(!t1hrs, "#,##0.00")
      txtT2hrs.Text = Format(!t2hrs, "#,##0.00")
      txtBrkhrs.Text = Format(!brkhrs, "#,##0.00")
      txtNitepremHrs.Text = Format(!nitepremhrs, "#,##0.00")
      chkRequired.Value = IIf(!Required = "Y", 1, 0)
    Else
      ClearText
    End If
  End With
  
End Sub

Private Sub txtBrkhrs_GotFocus()
  If IsDate(txtT1out.Text) And IsDate(txtT2in.Text) Then
    txtBrkhrs.ReadOnly = True
  Else
    txtBrkhrs.ReadOnly = False
  End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchRecord KeyAscii, txtSearch, rShift, txtSearch.Text, mSort
  End If
End Sub

Private Sub ClearText()

  txtShiftcode.Text = ""
  txtT1In.Text = ""
  txtT2in.Text = ""
  txtT1out.Text = ""
  txtT2out.Text = ""
  txtBrkstart.Text = ""
  txtBrkend.Text = ""
  txtNitePremstart.Text = ""
  txtNitePremend.Text = ""
  txtT1Hrs.Text = "0.00"
  txtT2hrs.Text = "0.00"
  txtBrkhrs.Text = "0.00"
  txtNitepremHrs.Text = "0.00"
  chkRequired.Value = 0
  
End Sub

Private Sub txtT1In_LostFocus()
  Compute_hours txtT1In, txtT1out, txtT1Hrs
End Sub

Private Sub txtT1Out_LostFocus()
  Compute_hours txtT1In, txtT1out, txtT1Hrs
  Compute_hours txtT1out, txtT2in, txtBrkhrs
End Sub

Private Sub txtT2in_KeyPress(KeyAscii As Integer)
    If CDbl(txtT1Hrs.Text) <= 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtT2In_LostFocus()
  Compute_hours txtT2in, txtT2out, txtT2hrs
  Compute_hours txtT1out, txtT2in, txtBrkhrs
End Sub

Private Sub txtT2out_KeyPress(KeyAscii As Integer)
    If CDbl(txtT1Hrs.Text) <= 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtT2out_LostFocus()
  Compute_hours txtT2in, txtT2out, txtT2hrs
End Sub

Private Sub txtNitePremstart_LostFocus()
  Compute_hours txtNitePremstart, txtNitePremend, txtNitepremHrs
End Sub

Private Sub txtNitePremend_LostFocus()
  Compute_hours txtNitePremstart, txtNitePremend, txtNitepremHrs
End Sub

Private Sub Compute_hours(ByRef objTin As Object, ByRef objTout As Object, ByRef objThrs As Object)

  Dim mTime         As Double
  
  If IsDate(objTin.Text) And IsDate(objTout.Text) Then
    If CDate(objTin.Text) > CDate(objTout.Text) Then
      mTime = 24 + Format(Round(DateDiff("N", objTin.Text, "12:00 am") / 60, 2), "#,##0.00")
      objThrs.Value = Format(Round(DateDiff("N", "12:00 am", objTout.Text) / 60, 2), "#,##0.00") + mTime
    Else
      objThrs.Value = Format(Round(DateDiff("N", objTin.Text, objTout.Text) / 60, 2), "#,##0.00")
      If objThrs.Value = 0 Then objThrs.Value = 24
    End If
  Else
    objThrs.Value = 0
  End If
End Sub








