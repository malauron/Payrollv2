VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDSSS 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8460
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab tabSSS 
      Height          =   5775
      Left            =   105
      TabIndex        =   9
      Top             =   570
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
      Caption         =   "Maintain SSS Table|View SSS Table"
      Align           =   0
      CurrTab         =   1
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
         Left            =   -8835
         TabIndex        =   10
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
         Begin VB.Frame frmeSSS 
            BackColor       =   &H00F6F8F8&
            Enabled         =   0   'False
            Height          =   2955
            Left            =   135
            TabIndex        =   24
            Top             =   180
            Width           =   4365
            Begin TDBNumber6Ctl.TDBNumber txtFrom 
               Height          =   300
               Left            =   2220
               TabIndex        =   3
               Top             =   585
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":0000
               Caption         =   "frmMDSSS.frx":0020
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":008C
               Keys            =   "frmMDSSS.frx":00AA
               Spin            =   "frmMDSSS.frx":00F4
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin TDBText6Ctl.TDBText txtSSSBCode 
               Height          =   300
               Left            =   2220
               TabIndex        =   2
               Top             =   210
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDSSS.frx":011C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":0188
               Key             =   "frmMDSSS.frx":01A6
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
            Begin TDBNumber6Ctl.TDBNumber txtTo 
               Height          =   300
               Left            =   2220
               TabIndex        =   4
               Top             =   960
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":01EA
               Caption         =   "frmMDSSS.frx":020A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":0276
               Keys            =   "frmMDSSS.frx":0294
               Spin            =   "frmMDSSS.frx":02DE
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin TDBNumber6Ctl.TDBNumber txtSalaryCredit 
               Height          =   300
               Left            =   2220
               TabIndex        =   5
               Top             =   1335
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":0306
               Caption         =   "frmMDSSS.frx":0326
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":0392
               Keys            =   "frmMDSSS.frx":03B0
               Spin            =   "frmMDSSS.frx":03FA
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin TDBNumber6Ctl.TDBNumber txtER 
               Height          =   300
               Left            =   2220
               TabIndex        =   6
               Top             =   1710
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":0422
               Caption         =   "frmMDSSS.frx":0442
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":04AE
               Keys            =   "frmMDSSS.frx":04CC
               Spin            =   "frmMDSSS.frx":0516
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin TDBNumber6Ctl.TDBNumber txtEE 
               Height          =   300
               Left            =   2220
               TabIndex        =   7
               Top             =   2085
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":053E
               Caption         =   "frmMDSSS.frx":055E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":05CA
               Keys            =   "frmMDSSS.frx":05E8
               Spin            =   "frmMDSSS.frx":0632
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin TDBNumber6Ctl.TDBNumber txtEC 
               Height          =   300
               Left            =   2220
               TabIndex        =   8
               Top             =   2475
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3519
               _ExtentY        =   529
               Calculator      =   "frmMDSSS.frx":065A
               Caption         =   "frmMDSSS.frx":067A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":06E6
               Keys            =   "frmMDSSS.frx":0704
               Spin            =   "frmMDSSS.frx":074E
               AlignHorizontal =   1
               AlignVertical   =   2
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
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "EC"
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
               TabIndex        =   31
               Top             =   2505
               Width           =   855
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Employee (EE) Share"
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
               Left            =   180
               TabIndex        =   30
               Top             =   2115
               Width           =   1920
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Employer (ER) Share"
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
               Left            =   255
               TabIndex        =   29
               Top             =   1725
               Width           =   1830
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "To"
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
               Left            =   525
               TabIndex        =   28
               Top             =   990
               Width           =   1560
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Salary Credit"
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
               Left            =   240
               TabIndex        =   27
               Top             =   1350
               Width           =   1845
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bracket Code"
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
               Left            =   525
               TabIndex        =   26
               Top             =   240
               Width           =   1560
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Amount:  From"
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
               Left            =   555
               TabIndex        =   25
               Top             =   615
               Width           =   1545
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerCity 
         Height          =   5460
         Left            =   15
         TabIndex        =   11
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
         Begin TrueOleDBGrid80.TDBGrid gridSSS 
            Height          =   4740
            Left            =   90
            TabIndex        =   1
            Top             =   600
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   8361
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Bracket Code"
            Columns(0).DataField=   "SSSBCode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "From"
            Columns(1).DataField=   "FromAmount"
            Columns(1).NumberFormat=   "#,#0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "To"
            Columns(2).DataField=   "ToAmount"
            Columns(2).NumberFormat=   "#,#0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Salary Credit"
            Columns(3).DataField=   "SalaryCredit"
            Columns(3).NumberFormat=   "#,#0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "ER Share"
            Columns(4).DataField=   "ER"
            Columns(4).NumberFormat=   "#,#0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "EE Share"
            Columns(5).DataField=   "EE"
            Columns(5).NumberFormat=   "#,#0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "EC"
            Columns(6).DataField=   "EC"
            Columns(6).NumberFormat=   "#,#0.00"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2037"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1958"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1852"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1773"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2037"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1958"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=1852"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1773"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1746"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1667"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
         Begin TDBText6Ctl.TDBText txtSearchBoxSSS 
            Height          =   300
            Left            =   1365
            TabIndex        =   0
            Top             =   210
            Width           =   6150
            _Version        =   65536
            _ExtentX        =   10848
            _ExtentY        =   529
            Caption         =   "frmMDSSS.frx":0776
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDSSS.frx":07E2
            Key             =   "frmMDSSS.frx":0800
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH"
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
            Left            =   330
            TabIndex        =   12
            Top             =   255
            Width           =   870
         End
      End
      Begin C1SizerLibCtl.C1Elastic SizerModeofPay 
         Height          =   5460
         Left            =   8865
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   510
            Width           =   6045
            Begin TDBText6Ctl.TDBText TDBText8 
               Height          =   300
               Left            =   1800
               TabIndex        =   15
               Top             =   225
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   529
               Caption         =   "frmMDSSS.frx":0844
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":08B0
               Key             =   "frmMDSSS.frx":08CE
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
               TabIndex        =   16
               Top             =   555
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDSSS.frx":0912
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":097E
               Key             =   "frmMDSSS.frx":099C
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
               TabIndex        =   17
               Top             =   885
               Width           =   4005
               _Version        =   65536
               _ExtentX        =   7056
               _ExtentY        =   529
               Caption         =   "frmMDSSS.frx":09E0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDSSS.frx":0A4C
               Key             =   "frmMDSSS.frx":0A6A
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
               TabIndex        =   20
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
               TabIndex        =   19
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
               TabIndex        =   18
               Top             =   945
               Width           =   990
            End
         End
         Begin TDBText6Ctl.TDBText TDBText11 
            Height          =   300
            Left            =   1980
            TabIndex        =   21
            Top             =   165
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7056
            _ExtentY        =   529
            Caption         =   "frmMDSSS.frx":0AAE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmMDSSS.frx":0B1A
            Key             =   "frmMDSSS.frx":0B38
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
            TabIndex        =   22
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
            FormatString    =   $"frmMDSSS.frx":0B7C
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
            TabIndex        =   23
            Top             =   240
            Width           =   915
         End
      End
   End
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   33
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
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
      TabIndex        =   39
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
Attribute VB_Name = "frmMDSSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Recordsets
Option Explicit
Dim SSS As ADODB.Recordset

'Booleans
Dim mAdd As Boolean
Dim mEdit As Boolean
Dim mTransActive As Boolean

'storage
Dim mCode As Integer
Dim mSSSSortField As String

Private Sub Form_Load()
    
    Add_MDIButton Me.Name, TitleBar.Caption
    
    Call LoadSSS
    
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
    frmeSSS.Enabled = True
    gridSSS.Enabled = False
    txtSearchBoxSSS.Enabled = False
    mCode = txtSSSBCode.Text
    tabSSS.CurrTab = 0
    SafeSetFocus txtFrom
    mEdit = True
Else
    If txtFrom.Value = 0 Then
        MsgBox "You need to specify the salary range.", vbInformation
        SafeSetFocus txtFrom
        Exit Sub
    End If
    
    If txtTo.Value = 0 Then
        MsgBox "You need to specify the salary range.", vbInformation
        SafeSetFocus txtTo
        Exit Sub
    End If
    
    If txtFrom.Value > txtTo.Value Then
        MsgBox "Invalid salary range.", vbInformation
        txtTo.Value = 0
        txtFrom.Value = 0
        SafeSetFocus txtFrom
        Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        ConMain.Execute "update sss set fromAmount = " & txtFrom.Value & ", toAmount = " & txtTo.Value & ", salarycredit = " & txtSalaryCredit.Value & ", er = " & txtSalaryCredit.Value & ", er = " & txtER.Value & ", ee =" & txtEE.Value & ", ec = " & txtEC.Value & " where sssbcode = '" & txtSSSBCode.Text & "'"
    ConMain.CommitTrans
    gridSSS.Enabled = True
    frmeSSS.Enabled = False
    txtSearchBoxSSS.Enabled = True
    SSS.Requery
    pointmetdg gridSSS, SSS, "sssbcode", mCode
    mEdit = False
    mTransActive = False
    cmdMenu(1).Caption = "&Edit"
    Lock_Button "TTTFTT", cmdMenu, 5
    tabSSS.CurrTab = 1
End If
End Sub

Private Sub Delete_Record()
If SSS.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
        gridSSS.Delete
    End If
End If
End Sub

Private Sub Cancel_Transaction()
If mAdd = True Then
    cmdMenu(0).Caption = "&New"
    If SSS.RecordCount > 0 Then
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
frmeSSS.Enabled = False
txtSearchBoxSSS.Enabled = True
gridSSS.Enabled = True
gridSSS_RowColChange gridSSS.Row, gridSSS.Col
tabSSS.CurrTab = 1
End Sub

Private Sub Print_Record()

End Sub

Private Sub Close_Form()
Unload Me
End Sub

Private Sub ClearFields()
    txtSSSBCode.Text = "AUTO GENERATED..."
    txtFrom.Value = 0
    txtTo.Value = 0
    txtSalaryCredit.Value = 0
    txtER.Value = 0
    txtEE.Value = 0
    txtEC.Value = 0
End Sub

Private Sub Add_Record()
If cmdMenu(0).Caption = "&New" Then
    mTransActive = True
    cmdMenu(0).Caption = "&Save"
    Lock_Button "TFFTFF", cmdMenu, 5
    frmeSSS.Enabled = True
    gridSSS.Enabled = False
    txtSearchBoxSSS.Enabled = False
    Call ClearFields
    tabSSS.CurrTab = 0
    SafeSetFocus txtFrom
    mAdd = True
Else
    If txtFrom.Value < 0 Then
        MsgBox "You need to specify the salary range.", vbInformation
        SafeSetFocus txtFrom
        Exit Sub
    End If
    
    If txtTo.Value < 0 Then
        MsgBox "You need to specify the salary range.", vbInformation
        SafeSetFocus txtTo
        Exit Sub
    End If
    
    If txtFrom.Value > txtTo.Value Then
        MsgBox "Invalid salary range.", vbInformation
        txtTo.Value = 0
        txtFrom.Value = 0
        SafeSetFocus txtFrom
        Exit Sub
    End If
        
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
        mTransActive = True
        txtSSSBCode.Text = LastCode("SSS")
        ConMain.Execute "insert into SSS values ('" & txtSSSBCode.Text & "', " & _
                              "" & txtTo.Value & ", " & txtFrom.Value & ", " & txtSalaryCredit.Value & ", " & txtER.Value & ", " & txtEE.Value & ", " & txtEC.Value & ")"
    ConMain.CommitTrans
    gridSSS.Enabled = True
    frmeSSS.Enabled = False
    txtSearchBoxSSS.Enabled = True
    mCode = txtSSSBCode.Text
    SSS.Requery
    pointmetdg gridSSS, SSS, "SSSBCode", mCode
    mAdd = False
    mTransActive = False
    cmdMenu(0).Caption = "&New"
    Lock_Button "TTTFTT", cmdMenu, 5
    tabSSS.CurrTab = 1
End If
End Sub

Private Sub LoadSSS()
DoEvents
NetOpen SSS, "select * from SSS order by sssbcode"
DoEvents
If SSS.State = adStateOpen Then
    If SSS.RecordCount > 0 Then
        SSS.MoveFirst
        Lock_Button "TTTFTT", cmdMenu, 5
    Else
        Lock_Button "TFFFTT", cmdMenu, 5
    End If
    Set gridSSS.DataSource = SSS
    mSSSSortField = "sssbcode"
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
    
    With tabSSS
        .Top = fraButton.Top + fraButton.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
    
    With gridSSS
        .Left = 150
        .Width = Me.ScaleWidth - 300
        .Height = tabSSS.Height - .Top - 1000
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmMDSSS = Nothing
End Sub

Sub FormCenter(Frm As Form)
    Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub

Private Sub gridSSS_HeadClick(ByVal ColIndex As Integer)
If SSS.RecordCount > 0 Then
    mSSSSortField = gridSSS.Columns(ColIndex).DataField
    SSS.Sort = mSSSSortField
End If
End Sub

Private Sub gridSSS_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With SSS
    If .RecordCount > 0 Then
        txtSSSBCode.Text = !sssbcode
        txtFrom.Value = !FromAmount
        txtTo.Value = !toAmount
        txtSalaryCredit.Value = !salarycredit
        txtER.Value = !er
        txtEE.Value = !ee
        txtEC.Value = !ec
    Else
        Call ClearFields
    End If
End With
End Sub

Private Sub tabSSS_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
If mAdd = True Then
    Cancel = 1
End If
If mEdit = True Then
    Cancel = 1
End If
End Sub

Private Sub txtEC_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtEE_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtER_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtFrom_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtSalaryCredit_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtSearchBoxSSS_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
Else
    SearchRecord Keyascii, txtSearchBoxSSS, SSS, txtSearchBoxSSS.Text, mSSSSortField
End If
End Sub

Private Sub txtSSSBCode_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtTo_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
