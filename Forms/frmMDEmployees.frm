VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDEmployees 
   Appearance      =   0  'Flat
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   9660
   ClientLeft      =   210
   ClientTop       =   60
   ClientWidth     =   15825
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
   Icon            =   "frmMDEmployees.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleMode       =   0  'User
   ScaleWidth      =   15825
   WindowState     =   2  'Maximized
   Begin LinkProPayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   60
      TabIndex        =   102
      Top             =   15
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   1
         Left            =   1230
         TabIndex        =   103
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   0
         Left            =   75
         TabIndex        =   104
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   2
         Left            =   2385
         TabIndex        =   105
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   3
         Left            =   3540
         TabIndex        =   106
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   4
         Left            =   4695
         TabIndex        =   107
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   420
         Index           =   5
         Left            =   5850
         TabIndex        =   108
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
      Left            =   7935
      TabIndex        =   91
      Top             =   75
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   609
      Caption         =   "Employee Master Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   4210752
   End
   Begin C1SizerLibCtl.C1Tab tabEmployee 
      Height          =   8805
      Left            =   -30
      TabIndex        =   60
      Top             =   615
      Width           =   15435
      _cx             =   27226
      _cy             =   15531
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
      Caption         =   "Employee Listing|Employee Details"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   8490
         Left            =   16650
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   300
         Width           =   15405
         _cx             =   27173
         _cy             =   14975
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
         Appearance      =   4
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   8490
         Left            =   16350
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   300
         Width           =   15405
         _cx             =   27173
         _cy             =   14975
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
         Appearance      =   4
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   8490
         Left            =   16050
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   300
         Width           =   15405
         _cx             =   27173
         _cy             =   14975
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
         Appearance      =   4
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   8490
         Left            =   15
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   300
         Width           =   15405
         _cx             =   27173
         _cy             =   14975
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
         FloodStyle      =   0
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Frame fra 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   8415
            Index           =   0
            Left            =   60
            TabIndex        =   92
            Top             =   60
            Width           =   2940
            Begin VB.CheckBox chkIsActive 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               Caption         =   "Active"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   75
               TabIndex        =   117
               Top             =   2565
               Width           =   975
            End
            Begin TDBDate6Ctl.TDBDate tdbBirthdate 
               Height          =   300
               Left            =   1260
               TabIndex        =   9
               Top             =   6315
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   529
               Calendar        =   "frmMDEmployees.frx":6852
               Caption         =   "frmMDEmployees.frx":6958
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":69BE
               Keys            =   "frmMDEmployees.frx":69DC
               Spin            =   "frmMDEmployees.frx":6A3A
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
               Text            =   "01/21/2008"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   39468
               CenturyMode     =   0
            End
            Begin TDBText6Ctl.TDBText txtEmpNo 
               Height          =   300
               Left            =   45
               TabIndex        =   2
               Top             =   3135
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2831
               _ExtentY        =   529
               Caption         =   "frmMDEmployees.frx":6A62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6ACE
               Key             =   "frmMDEmployees.frx":6AEC
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
               MaxLength       =   10
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
            Begin TDBText6Ctl.TDBText txtBioMetId 
               Height          =   300
               Left            =   45
               TabIndex        =   3
               Top             =   3660
               Width           =   1605
               _Version        =   65536
               _ExtentX        =   2831
               _ExtentY        =   529
               Caption         =   "frmMDEmployees.frx":6B30
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6B9C
               Key             =   "frmMDEmployees.frx":6BBA
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
               MaxLength       =   20
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
            Begin TDBText6Ctl.TDBText txtLastname 
               Height          =   300
               Left            =   60
               TabIndex        =   4
               Top             =   4155
               Width           =   2700
               _Version        =   65536
               _ExtentX        =   4762
               _ExtentY        =   529
               Caption         =   "frmMDEmployees.frx":6BFE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6C6A
               Key             =   "frmMDEmployees.frx":6C88
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
               MaxLength       =   20
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
            Begin TDBText6Ctl.TDBText txtFirstname 
               Height          =   315
               Left            =   60
               TabIndex        =   5
               Top             =   4650
               Width           =   2700
               _Version        =   65536
               _ExtentX        =   4762
               _ExtentY        =   556
               Caption         =   "frmMDEmployees.frx":6CCC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6D38
               Key             =   "frmMDEmployees.frx":6D56
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
               MaxLength       =   20
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
            Begin TDBText6Ctl.TDBText txtMiddleName 
               Height          =   315
               Left            =   75
               TabIndex        =   6
               Top             =   5190
               Width           =   2700
               _Version        =   65536
               _ExtentX        =   4762
               _ExtentY        =   556
               Caption         =   "frmMDEmployees.frx":6D9A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6E06
               Key             =   "frmMDEmployees.frx":6E24
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
               MaxLength       =   20
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
            Begin TDBText6Ctl.TDBText txtAge 
               Height          =   300
               Left            =   1260
               TabIndex        =   10
               Top             =   6645
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   529
               Caption         =   "frmMDEmployees.frx":6E68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":6ED4
               Key             =   "frmMDEmployees.frx":6EF2
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
               MaxLength       =   70
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
            Begin TrueOleDBList80.TDBCombo tdbGender 
               Height          =   345
               Left            =   1260
               TabIndex        =   7
               Tag             =   "Municipal"
               Top             =   5565
               Width           =   1500
               _ExtentX        =   2646
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
               _PropDict       =   $"frmMDEmployees.frx":6F36
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
            Begin TrueOleDBList80.TDBCombo tdbCivilStatus 
               Height          =   345
               Left            =   1260
               TabIndex        =   8
               Tag             =   "Municipal"
               Top             =   5940
               Width           =   1500
               _ExtentX        =   2646
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
               _PropDict       =   $"frmMDEmployees.frx":6FE0
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
            Begin VB.Image imgPhoto 
               Height          =   2475
               Left            =   195
               Stretch         =   -1  'True
               Top             =   45
               Width           =   2475
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status"
               Height          =   255
               Left            =   -90
               TabIndex        =   101
               Top             =   6000
               Width           =   1245
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Birthdate"
               Height          =   255
               Left            =   195
               TabIndex        =   100
               Top             =   6330
               Width           =   960
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
               Height          =   255
               Left            =   570
               TabIndex        =   99
               Top             =   5655
               Width           =   585
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               Height          =   255
               Left            =   585
               TabIndex        =   98
               Top             =   6675
               Width           =   585
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
               Height          =   270
               Left            =   -540
               TabIndex        =   97
               Top             =   4980
               Width           =   1695
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               Height          =   270
               Left            =   -255
               TabIndex        =   96
               Top             =   4440
               Width           =   1215
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               Height          =   255
               Left            =   -240
               TabIndex        =   95
               Top             =   3960
               Width           =   1200
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Biometric Employee ID"
               Height          =   255
               Left            =   -150
               TabIndex        =   94
               Top             =   3450
               Width           =   2175
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Employee No."
               Height          =   255
               Left            =   -45
               TabIndex        =   93
               Top             =   2925
               Width           =   1290
            End
         End
         Begin C1SizerLibCtl.C1Tab tabEmployeeInfo 
            Height          =   8385
            Left            =   3000
            TabIndex        =   66
            Top             =   45
            Width           =   11955
            _cx             =   21087
            _cy             =   14790
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
            Caption         =   "Personal Information|Shift|Deductions|Payroll History"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   4
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   -1  'True
            TabsPerPage     =   4
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   -1  'True
            MultiRowOffset  =   0
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   1
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic C1Elastic21 
               Height          =   8070
               Left            =   12570
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   300
               Width           =   11925
               _cx             =   21034
               _cy             =   14235
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
               Begin TDBText6Ctl.TDBText txtShiftcode 
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   90
                  Tag             =   "Check or CC number"
                  Top             =   765
                  Visible         =   0   'False
                  Width           =   1485
                  _Version        =   65536
                  _ExtentX        =   2619
                  _ExtentY        =   450
                  Caption         =   "frmMDEmployees.frx":708A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmMDEmployees.frx":70F6
                  Key             =   "frmMDEmployees.frx":7114
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
                  MaxLength       =   20
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
               Begin TrueOleDBGrid80.TDBDropDown tddShift 
                  Height          =   1605
                  Left            =   3390
                  TabIndex        =   89
                  Top             =   945
                  Width           =   3375
                  _ExtentX        =   5953
                  _ExtentY        =   2831
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "Shift Code"
                  Columns(0).DataField=   "shiftcode"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "Shift Description"
                  Columns(1).DataField=   "shiftdesc"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "t1in"
                  Columns(2).DataField=   "t1in"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "t1out"
                  Columns(3).DataField=   "t1out"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   0
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "t2in"
                  Columns(4).DataField=   "t2in"
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   0
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "t2out"
                  Columns(5).DataField=   "t2out"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   6
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).MarqueeStyle=   3
                  Splits(0).AllowRowSizing=   0   'False
                  Splits(0).RecordSelectors=   0   'False
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0)._GSX_SAVERECORDSELECTORS=   0
                  Splits(0).AlternatingRowStyle=   -1  'True
                  Splits(0).DividerColor=   13160660
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=6"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2910"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2805"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(12)=   "Column(2).Width=3281"
                  Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3175"
                  Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
                  Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(18)=   "Column(3).Width=3281"
                  Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3175"
                  Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
                  Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(24)=   "Column(4).Width=3281"
                  Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=3175"
                  Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
                  Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
                  Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(30)=   "Column(5).Width=3281"
                  Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=3175"
                  Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
                  Splits(0)._ColumnProps(34)=   "Column(5).Visible=0"
                  Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
                  Splits.Count    =   1
                  AllowRowSizing  =   0   'False
                  Appearance      =   0
                  BorderStyle     =   1
                  ColumnHeaders   =   0   'False
                  DataMode        =   0
                  DefColWidth     =   0
                  Enabled         =   -1  'True
                  HeadLines       =   1
                  RowDividerStyle =   2
                  LayoutName      =   ""
                  LayoutFileName  =   ""
                  LayoutURL       =   ""
                  EmptyRows       =   0   'False
                  ListField       =   ""
                  DataField       =   ""
                  IntegralHeight  =   0   'False
                  FetchRowStyle   =   0   'False
                  AlternatingRowStyle=   -1  'True
                  DataMember      =   ""
                  ColumnFooters   =   0   'False
                  FootLines       =   1
                  RowTracking     =   -1  'True
                  DeadAreaBackColor=   16185592
                  ValueTranslate  =   0   'False
                  ScrollTrack     =   -1  'True
                  RowDividerColor =   13160660
                  RowSubDividerColor=   13160660
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
                  _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
                  _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HEFB6A7&"
                  _StyleDefs(16)  =   ":id=8,.fgcolor=&H0&"
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
                  _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
                  _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
                  _StyleDefs(57)  =   "Named:id=33:Normal"
                  _StyleDefs(58)  =   ":id=33,.parent=0"
                  _StyleDefs(59)  =   "Named:id=34:Heading"
                  _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(61)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(62)  =   "Named:id=35:Footing"
                  _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(64)  =   "Named:id=36:Selected"
                  _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(66)  =   "Named:id=37:Caption"
                  _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(68)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H8000000E&"
                  _StyleDefs(70)  =   "Named:id=39:EvenRow"
                  _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(72)  =   "Named:id=40:OddRow"
                  _StyleDefs(73)  =   ":id=40,.parent=33"
                  _StyleDefs(74)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(75)  =   ":id=41,.parent=34"
                  _StyleDefs(76)  =   "Named:id=42:FilterBar"
                  _StyleDefs(77)  =   ":id=42,.parent=33"
               End
               Begin TrueOleDBGrid80.TDBGrid tdgShift 
                  Height          =   2550
                  Left            =   135
                  TabIndex        =   36
                  Top             =   390
                  Width           =   10545
                  _ExtentX        =   18600
                  _ExtentY        =   4498
                  _LayoutType     =   4
                  _RowHeight      =   16
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "Days"
                  Columns(0).DataField=   "day"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "Shiftcode"
                  Columns(1).DataField=   "shiftcode"
                  Columns(1).DropDown=   "tddShift"
                  Columns(1).DropDown.vt=   8
                  Columns(1).ExternalEditor=   "txtShiftcode"
                  Columns(1).ExternalEditor.vt=   8
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "1st Time In"
                  Columns(2).DataField=   "t1in"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "1st Time Out"
                  Columns(3).DataField=   "t1out"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   0
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "2nd Time In"
                  Columns(4).DataField=   "t2in"
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   0
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "2nd Time Out"
                  Columns(5).DataField=   "t2out"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   6
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
                  Splits(0)._ColumnProps(0)=   "Columns.Count=6"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2540"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2461"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(7)=   "Column(1).Width=4154"
                  Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4075"
                  Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(12)=   "Column(2).Width=1799"
                  Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1720"
                  Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8196"
                  Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(18)=   "Column(3).Width=1852"
                  Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1773"
                  Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=8196"
                  Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(24)=   "Column(4).Width=1667"
                  Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=1588"
                  Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
                  Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=8196"
                  Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
                  Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
                  Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
                  Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=8196"
                  Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   0
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  DefColWidth     =   0
                  HeadLines       =   1
                  FootLines       =   1
                  MultipleLines   =   0
                  EmptyRows       =   -1  'True
                  CellTipsWidth   =   0
                  DeadAreaBackColor=   16185592
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
                  _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HF6F8F8&"
                  _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&HF6F8F8&,.fgcolor=&H80000012&"
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
                  _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
                  _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
                  _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
                  _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=-1"
                  _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
                  _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
                  _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
                  _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
                  _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
                  _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
                  _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
                  _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
                  _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
                  _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
                  _StyleDefs(58)  =   "Named:id=33:Normal"
                  _StyleDefs(59)  =   ":id=33,.parent=0"
                  _StyleDefs(60)  =   "Named:id=34:Heading"
                  _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(62)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(63)  =   "Named:id=35:Footing"
                  _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(65)  =   "Named:id=36:Selected"
                  _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(67)  =   "Named:id=37:Caption"
                  _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(69)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
                  _StyleDefs(71)  =   "Named:id=39:EvenRow"
                  _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(73)  =   "Named:id=40:OddRow"
                  _StyleDefs(74)  =   ":id=40,.parent=33"
                  _StyleDefs(75)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(76)  =   ":id=41,.parent=34"
                  _StyleDefs(77)  =   "Named:id=42:FilterBar"
                  _StyleDefs(78)  =   ":id=42,.parent=33"
               End
               Begin VB.Frame fra 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame5"
                  ForeColor       =   &H80000008&
                  Height          =   360
                  Index           =   3
                  Left            =   2055
                  TabIndex        =   116
                  Top             =   45
                  Width           =   6585
                  Begin VB.CheckBox chkLogBased 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "Computation of wages will be based on his or her time logs."
                     ForeColor       =   &H80000008&
                     Height          =   315
                     Left            =   0
                     TabIndex        =   35
                     Top             =   0
                     Width           =   6255
                  End
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00F6F8F8&
                  Caption         =   "Shift Schedule"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   -555
                  TabIndex        =   69
                  Top             =   90
                  Width           =   2085
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic6 
               Height          =   8070
               Left            =   15
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   300
               Width           =   11925
               _cx             =   21034
               _cy             =   14235
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
               FloodColor      =   0
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
               Begin VB.Frame fra 
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame1"
                  Height          =   7065
                  Index           =   1
                  Left            =   90
                  TabIndex        =   70
                  Top             =   -585
                  Width           =   11505
                  Begin VB.CheckBox ChkRegular 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "Regular"
                     ForeColor       =   &H80000008&
                     Height          =   315
                     Left            =   6375
                     TabIndex        =   34
                     Top             =   4455
                     Visible         =   0   'False
                     Width           =   2925
                  End
                  Begin VB.CheckBox chkSalToBank 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "for PAYROLL CREDIT UPLOAD"
                     ForeColor       =   &H80000008&
                     Height          =   315
                     Left            =   6375
                     TabIndex        =   33
                     Top             =   4185
                     Width           =   2925
                  End
                  Begin TDBText6Ctl.TDBText txtHouseNo 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   11
                     Top             =   15
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":7158
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":71C4
                     Key             =   "frmMDEmployees.frx":71E2
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
                  Begin TDBText6Ctl.TDBText txtTelno 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   16
                     Top             =   1815
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":7226
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":7292
                     Key             =   "frmMDEmployees.frx":72B0
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
                     MaxLength       =   20
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
                  Begin TrueOleDBList80.TDBCombo tdbBarangay 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   15
                     Top             =   1335
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":72F4
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
                  Begin TrueOleDBList80.TDBCombo tdbMunicipal 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   14
                     Top             =   1005
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":739E
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
                  Begin TrueOleDBList80.TDBCombo tdbProvince 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   13
                     Top             =   675
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7448
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
                  Begin TDBText6Ctl.TDBText txtMobileno 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   17
                     Top             =   2145
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":74F2
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":755E
                     Key             =   "frmMDEmployees.frx":757C
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
                     MaxLength       =   20
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
                  Begin TDBText6Ctl.TDBText txtEmrgncyName 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   38
                     Top             =   5340
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":75C0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":762C
                     Key             =   "frmMDEmployees.frx":764A
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
                     MaxLength       =   100
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
                  Begin TDBText6Ctl.TDBText txtEmrgncyNo 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   58
                     Top             =   5670
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":768E
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":76FA
                     Key             =   "frmMDEmployees.frx":7718
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
                     MaxLength       =   20
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
                  Begin TDBText6Ctl.TDBText txtEmrgncyEmail 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   59
                     Top             =   6000
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":775C
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":77C8
                     Key             =   "frmMDEmployees.frx":77E6
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
                     MaxLength       =   20
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
                  Begin TDBText6Ctl.TDBText txtEmail 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   18
                     Top             =   2475
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":782A
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":7896
                     Key             =   "frmMDEmployees.frx":78B4
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
                     MaxLength       =   50
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
                  Begin TDBText6Ctl.TDBText txtStreet 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   12
                     Top             =   345
                     Width           =   7695
                     _Version        =   65536
                     _ExtentX        =   13573
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":78F8
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":7964
                     Key             =   "frmMDEmployees.frx":7982
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
                     MaxLength       =   70
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
                  Begin TrueOleDBList80.TDBCombo tdbDivision 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   20
                     Top             =   3285
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":79C6
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
                  Begin TrueOleDBList80.TDBCombo tdbBranch 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   19
                     Top             =   2955
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7A70
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
                  Begin TrueOleDBList80.TDBCombo tdbEmpStat 
                     Height          =   300
                     Left            =   6375
                     TabIndex        =   31
                     Top             =   3180
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7B1A
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
                  Begin TrueOleDBList80.TDBCombo tdbJob 
                     Bindings        =   "frmMDEmployees.frx":7BC4
                     DataMember      =   "tdbJob"
                     Height          =   300
                     Left            =   6375
                     TabIndex        =   30
                     Top             =   2850
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7BD5
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
                  Begin TrueOleDBList80.TDBCombo tdbCostCenter 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   21
                     Top             =   3615
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7C7F
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
                  Begin TrueOleDBList80.TDBCombo tdbSection 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   22
                     Top             =   3945
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7D29
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
                  Begin TDBNumber6Ctl.TDBNumber txtMonthly_Rate 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   25
                     Top             =   1350
                     Width           =   1515
                     _Version        =   65536
                     _ExtentX        =   2672
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployees.frx":7DD3
                     Caption         =   "frmMDEmployees.frx":7DF3
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":7E59
                     Keys            =   "frmMDEmployees.frx":7E77
                     Spin            =   "frmMDEmployees.frx":7EC1
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "###,###,###,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   -2147483640
                     Format          =   "###,###,###,##0.00"
                     HighlightText   =   0
                     MarginBottom    =   1
                     MarginLeft      =   1
                     MarginRight     =   1
                     MarginTop       =   1
                     MaxValue        =   99999
                     MinValue        =   -99999
                     MousePointer    =   0
                     MoveOnLRKey     =   0
                     NegativeColor   =   255
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     ReadOnly        =   0
                     Separator       =   ","
                     ShowContextMenu =   -1
                     ValueVT         =   1245189
                     Value           =   0
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TrueOleDBList80.TDBCombo tdbRateType 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   24
                     Top             =   1005
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":7EE9
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
                  Begin TDBText6Ctl.TDBText txtBankAcctNo 
                     Height          =   300
                     Left            =   6375
                     TabIndex        =   32
                     Top             =   3840
                     Width           =   2865
                     _Version        =   65536
                     _ExtentX        =   5054
                     _ExtentY        =   529
                     Caption         =   "frmMDEmployees.frx":7F93
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":7FFF
                     Key             =   "frmMDEmployees.frx":801D
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
                     MaxLength       =   20
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
                  Begin TDBNumber6Ctl.TDBNumber txtMealAllow 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   28
                     Top             =   2400
                     Width           =   1515
                     _Version        =   65536
                     _ExtentX        =   2672
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployees.frx":8061
                     Caption         =   "frmMDEmployees.frx":8081
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":80E7
                     Keys            =   "frmMDEmployees.frx":8105
                     Spin            =   "frmMDEmployees.frx":814F
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "###,###,###,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   -2147483640
                     Format          =   "###,###,###,##0.00"
                     HighlightText   =   0
                     MarginBottom    =   1
                     MarginLeft      =   1
                     MarginRight     =   1
                     MarginTop       =   1
                     MaxValue        =   99999
                     MinValue        =   -99999
                     MousePointer    =   0
                     MoveOnLRKey     =   0
                     NegativeColor   =   255
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     ReadOnly        =   0
                     Separator       =   ","
                     ShowContextMenu =   -1
                     ValueVT         =   1245189
                     Value           =   0
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtFixedEarnings 
                     Height          =   300
                     Left            =   9420
                     TabIndex        =   29
                     Top             =   2400
                     Width           =   1515
                     _Version        =   65536
                     _ExtentX        =   2672
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployees.frx":8177
                     Caption         =   "frmMDEmployees.frx":8197
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":81FD
                     Keys            =   "frmMDEmployees.frx":821B
                     Spin            =   "frmMDEmployees.frx":8265
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "###,###,###,##0.00"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   -2147483640
                     Format          =   "###,###,###,##0.00"
                     HighlightText   =   0
                     MarginBottom    =   1
                     MarginLeft      =   1
                     MarginRight     =   1
                     MarginTop       =   1
                     MaxValue        =   99999
                     MinValue        =   -99999
                     MousePointer    =   0
                     MoveOnLRKey     =   0
                     NegativeColor   =   255
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     ReadOnly        =   0
                     Separator       =   ","
                     ShowContextMenu =   -1
                     ValueVT         =   1245189
                     Value           =   0
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtDaily_Rate 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   26
                     Top             =   1695
                     Width           =   1515
                     _Version        =   65536
                     _ExtentX        =   2672
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployees.frx":828D
                     Caption         =   "frmMDEmployees.frx":82AD
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":8313
                     Keys            =   "frmMDEmployees.frx":8331
                     Spin            =   "frmMDEmployees.frx":837B
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "###,##0.0000000"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   -2147483640
                     Format          =   "###,##0.0000000"
                     HighlightText   =   0
                     MarginBottom    =   1
                     MarginLeft      =   1
                     MarginRight     =   1
                     MarginTop       =   1
                     MaxValue        =   99999
                     MinValue        =   0
                     MousePointer    =   0
                     MoveOnLRKey     =   0
                     NegativeColor   =   255
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     ReadOnly        =   0
                     Separator       =   ","
                     ShowContextMenu =   -1
                     ValueVT         =   1245189
                     Value           =   0
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TDBNumber6Ctl.TDBNumber txtHourly_Rate 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   27
                     Top             =   2040
                     Width           =   1515
                     _Version        =   65536
                     _ExtentX        =   2672
                     _ExtentY        =   529
                     Calculator      =   "frmMDEmployees.frx":83A3
                     Caption         =   "frmMDEmployees.frx":83C3
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":8429
                     Keys            =   "frmMDEmployees.frx":8447
                     Spin            =   "frmMDEmployees.frx":8491
                     AlignHorizontal =   1
                     AlignVertical   =   0
                     Appearance      =   0
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     BtnPositioning  =   0
                     ClipMode        =   0
                     ClearAction     =   0
                     DecimalPoint    =   "."
                     DisplayFormat   =   "###,##0.0000000"
                     EditMode        =   0
                     Enabled         =   -1
                     ErrorBeep       =   0
                     ForeColor       =   -2147483640
                     Format          =   "###,##0.0000000"
                     HighlightText   =   0
                     MarginBottom    =   1
                     MarginLeft      =   1
                     MarginRight     =   1
                     MarginTop       =   1
                     MaxValue        =   99999
                     MinValue        =   0
                     MousePointer    =   0
                     MoveOnLRKey     =   0
                     NegativeColor   =   255
                     OLEDragMode     =   0
                     OLEDropMode     =   0
                     ReadOnly        =   0
                     Separator       =   ","
                     ShowContextMenu =   -1
                     ValueVT         =   1245189
                     Value           =   0
                     MaxValueVT      =   5
                     MinValueVT      =   5
                  End
                  Begin TrueOleDBList80.TDBCombo tdbPayFrequency2 
                     Height          =   345
                     Left            =   5985
                     TabIndex        =   144
                     Tag             =   "Municipal"
                     Top             =   4770
                     Visible         =   0   'False
                     Width           =   1500
                     _ExtentX        =   2646
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
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
                     Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                     Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
                     Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
                     Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
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
                     _PropDict       =   $"frmMDEmployees.frx":84B9
                     _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                     _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                     _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
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
                  Begin TrueOleDBList80.TDBCombo tdbPayFrequency 
                     Height          =   300
                     Left            =   6360
                     TabIndex        =   23
                     Top             =   675
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":8563
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
                  Begin TDBDate6Ctl.TDBDate txtDateHired 
                     Height          =   300
                     Left            =   1545
                     TabIndex        =   146
                     Top             =   4455
                     Width           =   1500
                     _Version        =   65536
                     _ExtentX        =   2646
                     _ExtentY        =   529
                     Calendar        =   "frmMDEmployees.frx":860D
                     Caption         =   "frmMDEmployees.frx":8713
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     DropDown        =   "frmMDEmployees.frx":8779
                     Keys            =   "frmMDEmployees.frx":8797
                     Spin            =   "frmMDEmployees.frx":87F5
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
                     Text            =   "01/21/2008"
                     ValidateMode    =   0
                     ValueVT         =   7
                     Value           =   39468
                     CenturyMode     =   0
                  End
                  Begin TrueOleDBList80.TDBCombo tdbBank 
                     Height          =   300
                     Left            =   6375
                     TabIndex        =   149
                     Top             =   3510
                     Width           =   2865
                     _ExtentX        =   5054
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
                     _PropDict       =   $"frmMDEmployees.frx":881D
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
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bank "
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   5040
                     TabIndex        =   150
                     Top             =   3540
                     Width           =   1290
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Date Hired"
                     Height          =   255
                     Left            =   480
                     TabIndex        =   147
                     Top             =   4470
                     Width           =   960
                  End
                  Begin VB.Label Label33 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Pay Frequency"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   5025
                     TabIndex        =   145
                     Top             =   705
                     Width           =   1290
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hourly Rate"
                     Height          =   315
                     Index           =   35
                     Left            =   4815
                     TabIndex        =   143
                     Top             =   2070
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Daily Rate"
                     Height          =   315
                     Index           =   34
                     Left            =   4815
                     TabIndex        =   142
                     Top             =   1725
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Fixed Earnings"
                     Height          =   420
                     Index           =   33
                     Left            =   7875
                     TabIndex        =   140
                     Top             =   2385
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Meal Allowance/Day"
                     Height          =   465
                     Index           =   32
                     Left            =   4905
                     TabIndex        =   139
                     Top             =   2385
                     Width           =   1395
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Bank Account No."
                     Height          =   315
                     Index           =   23
                     Left            =   4500
                     TabIndex        =   115
                     Top             =   3855
                     Width           =   1815
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rate Type"
                     Height          =   315
                     Index           =   18
                     Left            =   4815
                     TabIndex        =   114
                     Top             =   1065
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Monthly Rate"
                     Height          =   315
                     Index           =   19
                     Left            =   4815
                     TabIndex        =   113
                     Top             =   1380
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Section"
                     Height          =   315
                     Index           =   24
                     Left            =   -390
                     TabIndex        =   112
                     Top             =   4005
                     Width           =   1815
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Employment Status"
                     Height          =   315
                     Index           =   17
                     Left            =   4500
                     TabIndex        =   111
                     Top             =   3225
                     Width           =   1815
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Job Title"
                     Height          =   315
                     Index           =   15
                     Left            =   4650
                     TabIndex        =   110
                     Top             =   2895
                     Width           =   1680
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Branch"
                     Height          =   315
                     Index           =   14
                     Left            =   -75
                     TabIndex        =   86
                     Top             =   3000
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Division"
                     Height          =   315
                     Index           =   5
                     Left            =   -75
                     TabIndex        =   85
                     Top             =   3330
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cost center"
                     Height          =   315
                     Index           =   0
                     Left            =   -75
                     TabIndex        =   84
                     Top             =   3660
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Street Address"
                     Height          =   315
                     Index           =   1
                     Left            =   -30
                     TabIndex        =   82
                     Top             =   390
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Mobile No."
                     Height          =   315
                     Index           =   3
                     Left            =   105
                     TabIndex        =   81
                     Top             =   2190
                     Width           =   1335
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Telephone No."
                     Height          =   315
                     Index           =   2
                     Left            =   150
                     TabIndex        =   80
                     Top             =   1890
                     Width           =   1290
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Barangay"
                     Height          =   315
                     Index           =   4
                     Left            =   -15
                     TabIndex        =   79
                     Top             =   1395
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "City/Municipality"
                     Height          =   315
                     Index           =   6
                     Left            =   -15
                     TabIndex        =   78
                     Top             =   1080
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Province"
                     Height          =   315
                     Index           =   8
                     Left            =   -30
                     TabIndex        =   77
                     Top             =   735
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "House/Bldg No."
                     Height          =   315
                     Index           =   7
                     Left            =   -30
                     TabIndex        =   76
                     Top             =   60
                     Width           =   1485
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Name"
                     Height          =   315
                     Index           =   9
                     Left            =   120
                     TabIndex        =   75
                     Top             =   5385
                     Width           =   1335
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H8000000A&
                     X1              =   30
                     X2              =   9030
                     Y1              =   5250
                     Y2              =   5250
                  End
                  Begin VB.Label Label4 
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "IN CASE OF EMERGENCY, CONTACT:"
                     Height          =   315
                     Index           =   10
                     Left            =   195
                     TabIndex        =   74
                     Top             =   5025
                     Width           =   3390
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Contact No.(s)"
                     Height          =   315
                     Index           =   11
                     Left            =   120
                     TabIndex        =   73
                     Top             =   5715
                     Width           =   1335
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Email Address"
                     Height          =   315
                     Index           =   12
                     Left            =   120
                     TabIndex        =   72
                     Top             =   6045
                     Width           =   1335
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000016&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Email Address"
                     Height          =   315
                     Index           =   13
                     Left            =   120
                     TabIndex        =   71
                     Top             =   2520
                     Width           =   1335
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic15 
               Height          =   8070
               Left            =   12870
               TabIndex        =   118
               TabStop         =   0   'False
               Top             =   300
               Width           =   11925
               _cx             =   21034
               _cy             =   14235
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
               Begin VB.Frame fra 
                  BackColor       =   &H00F6F8F8&
                  BorderStyle     =   0  'None
                  Height          =   7170
                  Index           =   2
                  Left            =   90
                  TabIndex        =   119
                  Top             =   45
                  Width           =   9660
                  Begin VB.Frame Frame4 
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "HDMF"
                     Height          =   1590
                     Left            =   0
                     TabIndex        =   132
                     Top             =   3450
                     Width           =   4875
                     Begin VB.OptionButton optHDMFFixed 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Fixed"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   50
                        Top             =   840
                        Width           =   1500
                     End
                     Begin VB.OptionButton optHDMFAuto 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Auto Deduct"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   49
                        Top             =   510
                        Width           =   1500
                     End
                     Begin TDBText6Ctl.TDBText txtHDMFNo 
                        Height          =   300
                        Left            =   1785
                        TabIndex        =   48
                        Top             =   225
                        Width           =   2880
                        _Version        =   65536
                        _ExtentX        =   5080
                        _ExtentY        =   529
                        Caption         =   "frmMDEmployees.frx":88C7
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8933
                        Key             =   "frmMDEmployees.frx":8951
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
                        MaxLength       =   20
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
                     Begin TDBNumber6Ctl.TDBNumber txtHdmfAmt 
                        Height          =   300
                        Left            =   3030
                        TabIndex        =   51
                        Top             =   825
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":8995
                        Caption         =   "frmMDEmployees.frx":89B5
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8A1B
                        Keys            =   "frmMDEmployees.frx":8A39
                        Spin            =   "frmMDEmployees.frx":8A83
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin TDBNumber6Ctl.TDBNumber txtHDMFEr 
                        Height          =   300
                        Left            =   3030
                        TabIndex        =   52
                        Top             =   1170
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":8AAB
                        Caption         =   "frmMDEmployees.frx":8ACB
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8B31
                        Keys            =   "frmMDEmployees.frx":8B4F
                        Spin            =   "frmMDEmployees.frx":8B99
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "HDMF No."
                        Height          =   315
                        Index           =   16
                        Left            =   255
                        TabIndex        =   135
                        Top             =   240
                        Width           =   1485
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "HDMF EE"
                        Height          =   315
                        Index           =   28
                        Left            =   2130
                        TabIndex        =   134
                        Top             =   870
                        Width           =   1260
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "HDMF ER"
                        Height          =   315
                        Index           =   29
                        Left            =   2130
                        TabIndex        =   133
                        Top             =   1215
                        Width           =   1260
                     End
                  End
                  Begin VB.Frame Frame3 
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "PhilHealth"
                     Height          =   1785
                     Left            =   0
                     TabIndex        =   128
                     Top             =   5100
                     Width           =   4875
                     Begin VB.OptionButton optPhilHAuto 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Auto Deduct"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   54
                        Top             =   630
                        Width           =   1500
                     End
                     Begin VB.OptionButton optPhilHFixed 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Fixed"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   55
                        Top             =   960
                        Width           =   1500
                     End
                     Begin TDBText6Ctl.TDBText txtPhilHNo 
                        Height          =   300
                        Left            =   1785
                        TabIndex        =   53
                        Top             =   345
                        Width           =   2880
                        _Version        =   65536
                        _ExtentX        =   5080
                        _ExtentY        =   529
                        Caption         =   "frmMDEmployees.frx":8BC1
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8C2D
                        Key             =   "frmMDEmployees.frx":8C4B
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
                        MaxLength       =   20
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
                     Begin TDBNumber6Ctl.TDBNumber txtPhilHAmt 
                        Height          =   300
                        Left            =   2985
                        TabIndex        =   56
                        Top             =   945
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":8C8F
                        Caption         =   "frmMDEmployees.frx":8CAF
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8D15
                        Keys            =   "frmMDEmployees.frx":8D33
                        Spin            =   "frmMDEmployees.frx":8D7D
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin TDBNumber6Ctl.TDBNumber txtPhilEr 
                        Height          =   300
                        Left            =   2985
                        TabIndex        =   57
                        Top             =   1290
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":8DA5
                        Caption         =   "frmMDEmployees.frx":8DC5
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8E2B
                        Keys            =   "frmMDEmployees.frx":8E49
                        Spin            =   "frmMDEmployees.frx":8E93
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "PhilHealth No."
                        Height          =   315
                        Index           =   21
                        Left            =   255
                        TabIndex        =   131
                        Top             =   360
                        Width           =   1485
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "PhilH EE"
                        Height          =   315
                        Index           =   30
                        Left            =   2085
                        TabIndex        =   130
                        Top             =   990
                        Width           =   1260
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "PhilH ER"
                        Height          =   315
                        Index           =   31
                        Left            =   2085
                        TabIndex        =   129
                        Top             =   1335
                        Width           =   1260
                     End
                  End
                  Begin VB.Frame Frame2 
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "SSS"
                     Height          =   1980
                     Left            =   0
                     TabIndex        =   123
                     Top             =   1455
                     Width           =   4875
                     Begin VB.OptionButton optSSSFixed 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Fixed"
                        Height          =   300
                        Left            =   210
                        TabIndex        =   44
                        Top             =   900
                        Width           =   1515
                     End
                     Begin VB.OptionButton optSSSAuto 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Auto Deduct"
                        Height          =   300
                        Left            =   210
                        TabIndex        =   43
                        Top             =   570
                        Width           =   1515
                     End
                     Begin TDBText6Ctl.TDBText txtSSSno 
                        Height          =   300
                        Left            =   1770
                        TabIndex        =   42
                        Top             =   240
                        Width           =   2865
                        _Version        =   65536
                        _ExtentX        =   5054
                        _ExtentY        =   529
                        Caption         =   "frmMDEmployees.frx":8EBB
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":8F27
                        Key             =   "frmMDEmployees.frx":8F45
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
                        MaxLength       =   20
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
                     Begin TDBNumber6Ctl.TDBNumber txtSSSAmt 
                        Height          =   300
                        Left            =   3030
                        TabIndex        =   45
                        Top             =   885
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":8F89
                        Caption         =   "frmMDEmployees.frx":8FA9
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":900F
                        Keys            =   "frmMDEmployees.frx":902D
                        Spin            =   "frmMDEmployees.frx":9077
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin TDBNumber6Ctl.TDBNumber txtSssEr 
                        Height          =   300
                        Left            =   3030
                        TabIndex        =   46
                        Top             =   1230
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":909F
                        Caption         =   "frmMDEmployees.frx":90BF
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":9125
                        Keys            =   "frmMDEmployees.frx":9143
                        Spin            =   "frmMDEmployees.frx":918D
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin TDBNumber6Ctl.TDBNumber txtSssEc 
                        Height          =   300
                        Left            =   3030
                        TabIndex        =   47
                        Top             =   1575
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":91B5
                        Caption         =   "frmMDEmployees.frx":91D5
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":923B
                        Keys            =   "frmMDEmployees.frx":9259
                        Spin            =   "frmMDEmployees.frx":92A3
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "SSS No."
                        Height          =   315
                        Index           =   20
                        Left            =   210
                        TabIndex        =   127
                        Top             =   270
                        Width           =   1260
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "SSS EE:"
                        Height          =   315
                        Index           =   25
                        Left            =   2130
                        TabIndex        =   126
                        Top             =   930
                        Width           =   1260
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "SSS ER:"
                        Height          =   315
                        Index           =   26
                        Left            =   2130
                        TabIndex        =   125
                        Top             =   1275
                        Width           =   1260
                     End
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "SSS EC:"
                        Height          =   315
                        Index           =   27
                        Left            =   2130
                        TabIndex        =   124
                        Top             =   1620
                        Width           =   1260
                     End
                  End
                  Begin VB.Frame Frame1 
                     BackColor       =   &H00F6F8F8&
                     Caption         =   "Witholding Tax"
                     Height          =   1335
                     Left            =   0
                     TabIndex        =   120
                     Top             =   105
                     Width           =   4890
                     Begin VB.OptionButton optTaxFixed 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Fixed"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   40
                        Top             =   960
                        Width           =   1500
                     End
                     Begin VB.OptionButton optTaxAuto 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00F6F8F8&
                        Caption         =   "Auto Deduct"
                        Height          =   300
                        Left            =   255
                        TabIndex        =   121
                        Top             =   630
                        Width           =   1500
                     End
                     Begin TDBNumber6Ctl.TDBNumber txtTaxAmt 
                        Height          =   300
                        Left            =   1815
                        TabIndex        =   41
                        Top             =   945
                        Width           =   1440
                        _Version        =   65536
                        _ExtentX        =   2540
                        _ExtentY        =   529
                        Calculator      =   "frmMDEmployees.frx":92CB
                        Caption         =   "frmMDEmployees.frx":92EB
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":9351
                        Keys            =   "frmMDEmployees.frx":936F
                        Spin            =   "frmMDEmployees.frx":93B9
                        AlignHorizontal =   1
                        AlignVertical   =   0
                        Appearance      =   0
                        BackColor       =   -2147483643
                        BorderStyle     =   1
                        BtnPositioning  =   0
                        ClipMode        =   0
                        ClearAction     =   0
                        DecimalPoint    =   "."
                        DisplayFormat   =   "#,##0.00"
                        EditMode        =   0
                        Enabled         =   -1
                        ErrorBeep       =   0
                        ForeColor       =   -2147483640
                        Format          =   "#,##0.00"
                        HighlightText   =   0
                        MarginBottom    =   1
                        MarginLeft      =   1
                        MarginRight     =   1
                        MarginTop       =   1
                        MaxValue        =   99999
                        MinValue        =   -99999
                        MousePointer    =   0
                        MoveOnLRKey     =   0
                        NegativeColor   =   255
                        OLEDragMode     =   0
                        OLEDropMode     =   0
                        ReadOnly        =   0
                        Separator       =   ","
                        ShowContextMenu =   -1
                        ValueVT         =   1245189
                        Value           =   0
                        MaxValueVT      =   5
                        MinValueVT      =   5
                     End
                     Begin TDBText6Ctl.TDBText txtTinno 
                        Height          =   300
                        Left            =   1815
                        TabIndex        =   37
                        Top             =   285
                        Width           =   2865
                        _Version        =   65536
                        _ExtentX        =   5054
                        _ExtentY        =   529
                        Caption         =   "frmMDEmployees.frx":93E1
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DropDown        =   "frmMDEmployees.frx":944D
                        Key             =   "frmMDEmployees.frx":946B
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
                        MaxLength       =   20
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
                     Begin TrueOleDBList80.TDBCombo tdbWT 
                        Bindings        =   "frmMDEmployees.frx":94AF
                        DataMember      =   "tdbWT"
                        Height          =   300
                        Left            =   1815
                        TabIndex        =   39
                        Top             =   615
                        Width           =   2865
                        _ExtentX        =   5054
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
                        _PropDict       =   $"frmMDEmployees.frx":94BF
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
                     Begin VB.Label Label4 
                        BackColor       =   &H80000016&
                        BackStyle       =   0  'Transparent
                        Caption         =   "T.I.N. No."
                        Height          =   315
                        Index           =   22
                        Left            =   255
                        TabIndex        =   122
                        Top             =   300
                        Width           =   1095
                     End
                  End
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid15 
                  Height          =   2205
                  Left            =   1410
                  TabIndex        =   136
                  Top             =   7380
                  Visible         =   0   'False
                  Width           =   9165
                  _cx             =   16166
                  _cy             =   3889
                  Appearance      =   3
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
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
                  BackColorBkg    =   -2147483626
                  BackColorAlternate=   13431287
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmMDEmployees.frx":9569
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
                  ExplorerBar     =   1
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
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid16 
                  Height          =   2520
                  Left            =   555
                  TabIndex        =   137
                  Top             =   7365
                  Visible         =   0   'False
                  Width           =   9165
                  _cx             =   16166
                  _cy             =   4445
                  Appearance      =   3
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
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
                  BackColorBkg    =   -2147483626
                  BackColorAlternate=   13431287
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   1
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
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmMDEmployees.frx":9618
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
                  ExplorerBar     =   1
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
               Begin VB.Label Label21 
                  BackColor       =   &H80000016&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Deduction Summary:"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   645
                  TabIndex        =   138
                  Top             =   6420
                  Visible         =   0   'False
                  Width           =   2100
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   8070
               Left            =   13170
               TabIndex        =   148
               TabStop         =   0   'False
               Top             =   300
               Width           =   11925
               _cx             =   21034
               _cy             =   14235
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
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8490
         Left            =   -16020
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   300
         Width           =   15405
         _cx             =   27173
         _cy             =   14975
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
         Begin TrueOleDBGrid80.TDBGrid tdgEmployee 
            Height          =   3750
            Left            =   240
            TabIndex        =   141
            Top             =   1005
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   6615
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Employee Number"
            Columns(0).DataField=   "dummycode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Name"
            Columns(1).DataField=   "fullname"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3096"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3016"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7938"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7858"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid80.TDBGrid tdgEmployees 
            Height          =   4260
            Left            =   690
            TabIndex        =   83
            Top             =   6030
            Visible         =   0   'False
            Width           =   13950
            _ExtentX        =   24606
            _ExtentY        =   7514
            _LayoutType     =   4
            _RowHeight      =   16
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Employee ID"
            Columns(0).DataField=   "empno"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Name"
            Columns(1).DataField=   "fullname"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Division"
            Columns(2).DataField=   "division"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cost Center"
            Columns(3).DataField=   "costcenter"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Status"
            Columns(4).DataField=   "civilstatus"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2328"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2249"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=8890"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=8811"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=4207"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4128"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=4260"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4180"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
            _StyleDefs(54)  =   "Named:id=33:Normal"
            _StyleDefs(55)  =   ":id=33,.parent=0"
            _StyleDefs(56)  =   "Named:id=34:Heading"
            _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   ":id=34,.wraptext=-1"
            _StyleDefs(59)  =   "Named:id=35:Footing"
            _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(61)  =   "Named:id=36:Selected"
            _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(63)  =   "Named:id=37:Caption"
            _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(65)  =   "Named:id=38:HighlightRow"
            _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
            _StyleDefs(67)  =   "Named:id=39:EvenRow"
            _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(69)  =   "Named:id=40:OddRow"
            _StyleDefs(70)  =   ":id=40,.parent=33"
            _StyleDefs(71)  =   "Named:id=41:RecordSelector"
            _StyleDefs(72)  =   ":id=41,.parent=34"
            _StyleDefs(73)  =   "Named:id=42:FilterBar"
            _StyleDefs(74)  =   ":id=42,.parent=33"
         End
         Begin VB.Frame fraSearch 
            BackColor       =   &H00F6F8F8&
            Height          =   720
            Left            =   165
            TabIndex        =   87
            Top             =   105
            Width           =   12285
            Begin TDBText6Ctl.TDBText txtSearch 
               Height          =   315
               Left            =   1500
               TabIndex        =   0
               Top             =   255
               Width           =   6180
               _Version        =   65536
               _ExtentX        =   10901
               _ExtentY        =   556
               Caption         =   "frmMDEmployees.frx":9683
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmMDEmployees.frx":96EF
               Key             =   "frmMDEmployees.frx":970D
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
            Begin TrueOleDBList80.TDBCombo tdbSort 
               Height          =   315
               Left            =   8385
               TabIndex        =   1
               Top             =   255
               Width           =   3060
               _ExtentX        =   5398
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
               Columns(0).DataField=   "code"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Description"
               Columns(1).DataField=   "description"
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
               _PropDict       =   $"frmMDEmployees.frx":9751
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
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000016&
               BackStyle       =   0  'Transparent
               Caption         =   "Sort"
               Height          =   195
               Left            =   7470
               TabIndex        =   109
               Top             =   315
               Width           =   765
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "SEARCH"
               Height          =   255
               Left            =   375
               TabIndex        =   88
               Top             =   315
               Width           =   915
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgBrowsePic 
      Left            =   10905
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMDEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Dim rsEmployee      As ADODB.Recordset
Dim rsEmpShift      As ADODB.Recordset
Dim rsEmpShiftTmp   As ADODB.Recordset
Dim rsImages        As ADODB.Recordset
Dim mEmpImage       As ADODB.Stream
Dim mSort           As String
Dim mPicName        As String

Private Sub cmdEmployee_Click(Index As Integer)
  
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
  
    Dim rsTmp         As ADODB.Recordset
    Dim i             As Integer
    
    Add_MDIButton Me.Name, TitleBar.Caption
    
    tabEmployee.CurrTab = 0
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 0 To 1
        .AddNew
        Select Case i
          Case 0:   .Fields("code") = "dummycode"
                    .Fields("description") = "Employee Number"
          Case 1:   .Fields("code") = "fullname"
                    .Fields("description") = "Fullname"
        End Select
        .Update
      Next
    End With
    
    With tdbSort
        .BoundColumn = "CODE"
        .ListField = "Description"
        .Columns(0).DataField = "CODE"
        .Columns(1).DataField = "Description"
        .RowSource = rsTmp
        .BoundText = "fullname"
        mSort = "fullname"
        DoEvents
    End With
    
    Set rsTmp = Nothing

    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 2
        .AddNew
        .Fields("code") = i
        Select Case i
          Case 1: .Fields("description") = "Male"
          Case 2: .Fields("description") = "Female"
        End Select
        .Update
      Next
    End With
    
    With tdbGender
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
          Case 1: .Fields("description") = "Single"
          Case 2: .Fields("description") = "Married"
          Case 3: .Fields("description") = "Widow"
          Case 4: .Fields("description") = "Widower"
          Case 5: .Fields("description") = "Separated"
        End Select
        .Update
      Next
    End With
    
    With tdbCivilStatus
     .BoundColumn = "CODE"
     .ListField = "Description"
     .Columns(0).DataField = "CODE"
     .Columns(1).DataField = "Description"
     .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing

    NetOpen rsEmployee, "select x1.*,concat(x1.lastname , ', ' , x1.firstname ,' ', x1.middlename) as fullname,x2.division,x3.costcenter from employee x1 " & _
                            "left outer join division x2 on x1.divisioncode = x2.divisioncode " & _
                            "left outer join costcenter x3 on x1.costcentercode = x3.costcentercode order by x1.employeecode"
                            
    If rsEmployee.RecordCount > 0 Then
      rsEmployee.MoveFirst
      rsEmployee.Sort = "fullname"
    End If
    
    Set tdgEmployee.DataSource = rsEmployee
    
    bind_tdb ConMain, tdbPayFrequency, "select payfreqcode,payfreqname from payfrequency order by payfreqname", "payfreqname", "payfreqcode"
    bind_tdb ConMain, tdbRateType, "select ratetypecode, ratetypename from ratetypes order by ratetypename", "ratetypename", "ratetypecode"
    bind_tdb ConMain, tdbJob, "select jobtitlecode,jobtitle from jobtitle order by jobtitle", "jobtitle", "jobtitlecode"
    bind_tdb ConMain, tdbWT, "select wtcode, description from wt order by description", "description", "wtcode"
    bind_tdb ConMain, tdbEmpStat, "select empstatcode,empstatname from employmentstatus order by empstatname", "empstatname", "empstatcode"
    bind_tdb ConMain, tdbBank, "select bankcode,bankname from bank order by bankname", "bankname", "bankcode"
    
    CreateEmpShiftTmp
      
    cmdEmployee_Click 3

End Sub

Private Sub AddSave_Button_Clicked()
    
  Dim rsChk         As ADODB.Recordset
  Dim rsEmpPics     As ADODB.Recordset
  
  Dim mPhoto        As ADODB.Stream
  
  Dim mWT           As String
  Dim mProvince     As String
  Dim mMunicipal    As String
  Dim mBarangay     As String
  Dim mSection      As String
  Dim mBank         As String
  
  Dim mSSSAuto      As Integer
  Dim mPHILHAuto    As Integer
  Dim mHDMFAuto     As Integer
  Dim mTAXAuto      As Integer
  
  Dim mSSSAmt       As Double
  Dim mSSSEr        As Double
  Dim mSSSEc        As Double
  Dim mPHILHAmt     As Double
  Dim mPHILHEr      As Double
  Dim mHDMFAmt      As Double
  Dim mHdmfEr       As Double
  Dim mTAXAmt       As Double
  
  If cmdEmployee(0).Caption = "&New" Then
  
    Lock_Button "TFFTFF", cmdEmployee, 5
    cmdEmployee(0).Caption = "&Save"
    tabEmployee.CurrTab = 1
    tabEmployeeInfo.CurrTab = 0
    ClearText
    txtEmpNo.Text = "...."
    tdgShift.Columns("shiftcode").Locked = False
    Lock_Tab "FT", tabEmployee, 1
    Lock_Frame "TTTT", fra, 3
    Load_EmpShift
    bind_tdb ConMain, tdbProvince, "select provcode,provname from province order by provname", "provname", "provcode"
    bind_tdb ConMain, tdbBranch, "select branchcode,branch from branch order by branch", "branch", "branchcode"
    
  Else
  
    If Trim(txtLastname.Text) = "" Then
      MsgBox "Lastname is blank.", vbExclamation + vbOKOnly
      txtLastname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtFirstname.Text) = "" Then
      MsgBox "Firstname is blank.", vbExclamation + vbOKOnly
      txtFirstname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
      MsgBox "Middlename is blank.", vbExclamation + vbOKOnly
      txtMiddleName.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbGender.Text) = "" Or IsNull(tdbGender.SelectedItem) Or tdbGender.ApproxCount = 0 Then
      MsgBox "Please select a gender.", vbExclamation + vbOKOnly
      tdbGender.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbCivilStatus.Text) = "" Or IsNull(tdbCivilStatus.SelectedItem) Or tdbCivilStatus.ApproxCount = 0 Then
      MsgBox "Please select a civil status.", vbExclamation + vbOKOnly
      tdbCivilStatus.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(tdbBirthdate.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      tdbBirthdate.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtDateHired.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtDateHired.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbPayFrequency.Text) = "" And IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
        MsgBox "Please assign a payroll frequency.", vbExclamation + vbOKOnly
        tdbPayFrequency.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbProvince.Text) = "" Or IsNull(tdbProvince.SelectedItem) Or tdbProvince.ApproxCount = 0 Then
        mProvince = ""
    Else
        mProvince = tdbProvince.BoundText
    End If
    
    If Trim(tdbMunicipal.Text) = "" Or IsNull(tdbMunicipal.SelectedItem) Or tdbMunicipal.ApproxCount = 0 Then
        mMunicipal = ""
    Else
        mMunicipal = tdbMunicipal.BoundText
    End If
    
    If Trim(tdbBarangay.Text) = "" Or IsNull(tdbBarangay.SelectedItem) Or tdbBarangay.ApproxCount = 0 Then
        mBarangay = ""
    Else
        mBarangay = tdbBarangay.BoundText
    End If
    
    If Trim(tdbBank.Text) = "" Or IsNull(tdbBank.SelectedItem) Or tdbBank.ApproxCount <= 0 Then
      mBank = "Null"
    Else
      mBank = tdbBank.BoundText
    End If
    
    If Trim(tdbBranch.Text) = "" Or IsNull(tdbBranch.SelectedItem) Or tdbBranch.ApproxCount = 0 Then
      MsgBox "Please select a branch.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbBranch.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbDivision.Text) = "" Or IsNull(tdbDivision.SelectedItem) Or tdbDivision.ApproxCount = 0 Then
      MsgBox "Please select a division.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbDivision.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbCostCenter.Text) = "" Or IsNull(tdbCostCenter.SelectedItem) Or tdbDivision.ApproxCount = 0 Then
      MsgBox "Please select a cost center.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbCostCenter.SetFocus
      Exit Sub
    End If
    
    
    If Trim(tdbRateType.Text) = "" Or IsNull(tdbRateType.SelectedItem) Or tdbRateType.ApproxCount = 0 Then
        MsgBox "Please select a rate type.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbRateType.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbSection.Text) <> "" And Not IsNull(tdbSection.SelectedItem) And tdbSection.ApproxCount > 0 Then
        mSection = tdbSection.BoundText
    Else
        mSection = "Null"
    End If
    
    If Trim(tdbJob.Text) = "" Or IsNull(tdbJob.SelectedItem) Or tdbJob.ApproxCount = 0 Then
        MsgBox "Please select a job description.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbJob.SetFocus
        Exit Sub
    End If
        
    If Trim(tdbEmpStat.Text) = "" Or IsNull(tdbEmpStat.SelectedItem) Or tdbEmpStat.ApproxCount = 0 Then
        MsgBox "Please select an employment status.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbEmpStat.SetFocus
        Exit Sub
    End If
    
    
    
    
    If optSSSAuto.Value = True Then
        mSSSAuto = 1
        mSSSAmt = 0
        mSSSEr = 0
        mSSSEc = 0
    Else
        mSSSAuto = 0
        mSSSAmt = Format(txtSSSAmt.Text, "###0.00")
        mSSSEr = Format(txtSssEr.Text, "###0.00")
        mSSSEc = Format(txtSssEc.Text, "###0.00")
    End If
    
    If optPhilHAuto.Value = True Then
        mPHILHAuto = 1
        mPHILHAmt = 0
        mPHILHEr = 0
    Else
        mPHILHAuto = 0
        mPHILHAmt = Format(txtPhilHAmt.Text, "###0.00")
        mPHILHEr = Format(txtPhilEr.Text, "###0.00")
    End If

    If optHDMFAuto.Value = True Then
        mHDMFAuto = 1
    Else
        mHDMFAuto = 0
        mHDMFAmt = Format(txtHdmfAmt.Text, "###0.00")
        mHdmfEr = Format(txtHDMFEr.Text, "###0.00")
    End If

    If optTaxAuto.Value = True Then
        mTAXAuto = 1
        mTAXAmt = 0
        If Trim(tdbWT.Text) = "" Or IsNull(tdbWT.SelectedItem) Or tdbWT.ApproxCount = 0 Then
            mWT = ""
        Else
            mWT = tdbWT.BoundText
        End If
    Else
        mTAXAuto = 0
        mTAXAmt = Format(txtTaxAmt.Text, "###0.00")
        mWT = ""
    End If

    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    txtEmpNo.Text = LastCode("Employee")
    
    If dlgBrowsePic.FileName <> "" Then
    
      ConMain.Execute "delete from emppics where employeecode = '" & txtEmpNo.Text & "'"
      
      Set mPhoto = New ADODB.Stream
      mPhoto.Type = adTypeBinary
      mPhoto.Open
      mPhoto.LoadFromFile (dlgBrowsePic.FileName)
      
      NetOpen rsEmpPics, "select * from emppics limit 0"
      With rsEmpPics
        .AddNew
        .Fields("employeecode") = txtEmpNo.Text
        .Fields("images") = mPhoto.Read
        .Fields("filename") = txtEmpNo.Text & "." & GetFileExt(dlgBrowsePic.FileTitle)
        .Update
        mPicName = .Fields("filename")
      End With
      Set rsEmpPics = Nothing
      
    End If
    
    ConMain.Execute "insert into employee (employeecode,dummycode,biometid,lastname,firstname,middlename,gender,civilstatus,birthdate,datehired, " & _
                          "houseno,street,provcode,muncode,brgycode,branchcode,divisioncode,costcentercode,telno,mobileno,email, " & _
                          "emrgncyname,emrgncyno,emrgncyemail,payfreqcode,filename, " & _
                          "wtcode,jobtitlecode,empstatcode,ratetypecode,monthly_rate,daily_rate,hourly_rate, " & _
                          "sssno,philhno,tinno,hdmfno,bankcode,bankacctno, " & _
                          "sectioncode,sssamt,ssser,sssec,philhamt,philher,taxamt,hdmfamt,hdmfer, " & _
                          "sssauto,philhauto,taxauto,hdmfauto,saltobank," & _
                          "regular,isactive,logbased,mealallow,fixedEarnings) values " & _
                          "(" & txtEmpNo.Text & ",'" & Format(txtEmpNo.Text, "00000000") & "','" & txtEmpNo.Text & "','" & UCase(txtLastname.Text) & "','" & UCase(txtFirstname.Text) & "','" & UCase(txtMiddleName.Text) & "','" & tdbGender.Text & "','" & tdbCivilStatus.Text & "','" & Format(tdbBirthdate.Text, "YYYY-MM-DD") & "','" & Format(txtDateHired.Text, "YYYY-MM-DD") & "', " & _
                          "'" & txtHouseNo.Text & "','" & txtStreet.Text & "','" & mProvince & "','" & mMunicipal & "','" & mBarangay & "','" & tdbBranch.BoundText & "','" & tdbDivision.BoundText & "','" & tdbCostCenter.BoundText & "','" & txtTelno.Text & "','" & txtMobileno.Text & "','" & txtEmail.Text & "', " & _
                          "'" & txtEmrgncyName.Text & "','" & txtEmrgncyNo.Text & "','" & txtEmrgncyEmail.Text & "','" & tdbPayFrequency.BoundText & "','" & mPicName & "', " & _
                          "'" & mWT & "','" & tdbJob.BoundText & "','" & tdbEmpStat.BoundText & "','" & tdbRateType.BoundText & "'," & Format(txtMonthly_Rate.Text, "##0.00") & "," & Format(txtDaily_Rate.Text, "##0.0000000") & "," & Format(txtHourly_Rate.Text, "##0.0000000") & ", " & _
                          "'" & txtSSSno.Text & "','" & txtPhilHNo.Text & "','" & txtTinno.Text & "','" & txtHDMFNo.Text & "'," & mBank & ",'" & txtBankAcctNo.Text & "', " & _
                          "" & mSection & "," & mSSSAmt & "," & mSSSEr & "," & mSSSEc & "," & mPHILHAmt & "," & mPHILHEr & "," & mTAXAmt & "," & mHDMFAmt & "," & mHdmfEr & ", " & _
                          "" & mSSSAuto & "," & mPHILHAuto & "," & mTAXAuto & "," & mHDMFAuto & ", '" & IIf(chkSalToBank.Value <> 0, "Y", "N") & "'," & _
                          "'" & IIf(ChkRegular.Value <> 0, "Y", "N") & "', '" & IIf(chkIsActive.Value <> 0, "Y", "N") & "', '" & IIf(chkLogBased.Value <> 0, "Y", "N") & "'," & Format(txtMealAllow.Text, "##0.00") & "," & Format(txtFixedEarnings.Text, "##0.00") & ")"
    
    rsEmpShiftTmp.MoveFirst
    Do While Not rsEmpShiftTmp.EOF
      ConMain.Execute "insert into empshift(employeecode,dayno,day,shiftcode) values " & _
                            "('" & txtEmpNo.Text & "','" & rsEmpShiftTmp!dayno & "','" & rsEmpShiftTmp!Day & "', " & _
                            "'" & rsEmpShiftTmp!shiftcode & "')"
      rsEmpShiftTmp.MoveNext
    Loop
                          
    ConMain.CommitTrans

    rsEmployee.Requery
    rsEmployee.Find "employeecode = '" & txtEmpNo.Text & "'"
        
    txtEmpNo.Text = Format(txtEmpNo.Text, "00000000")
    dlgBrowsePic.FileName = ""
    
    cmdEmployee_Click 3
      
  End If
  
End Sub

Private Sub EditUpdate_Button_Clicked()

  Dim rsChk         As ADODB.Recordset
  Dim rsEmpPics     As ADODB.Recordset
  
  Dim mPhoto        As ADODB.Stream
  
  Dim mWT           As String
  Dim mProvince     As String
  Dim mMunicipal    As String
  Dim mBarangay     As String
  Dim mSection      As String
  Dim mEmpno        As String
  Dim mBank         As String
  
  Dim mSSSAuto      As Integer
  Dim mPHILHAuto    As Integer
  Dim mHDMFAuto     As Integer
  Dim mTAXAuto      As Integer
  
  Dim mSSSAmt       As Double
  Dim mSSSEr        As Double
  Dim mSSSEc        As Double
  Dim mPHILHAmt     As Double
  Dim mPHILHEr      As Double
  Dim mHDMFAmt      As Double
  Dim mHdmfEr       As Double
  Dim mTAXAmt       As Double

  
  If cmdEmployee(1).Caption = "&Edit" Then
    
    Lock_Button "FTFTFF", cmdEmployee, 5
    tabEmployee.CurrTab = 1
    cmdEmployee(1).Caption = "&Update"
    Lock_Tab "FT", tabEmployee, 1
    Lock_Frame "TTTT", fra, 3
    tdgShift.Columns("shiftcode").Locked = False
    tabEmployeeInfo_Switch 0, 1, False
    Load_Address
    Load_Deductions
    Load_EmpShift
      
  Else
  
    If Trim(txtBioMetId.Text) <> rsEmployee!biometid Then
        NetOpen rsChk, "select biometid from employee where biometid = '" & Trim(txtBioMetId.Text) & "'"
        If rsChk.RecordCount > 0 Then
          MsgBox "Biometric Number already taken.", vbExclamation + vbOKOnly
          txtBioMetId.SetFocus
          Exit Sub
        End If
    End If
    
    Set rsChk = Nothing
  
    If Trim(txtLastname.Text) = "" Then
      MsgBox "Lastname is blank.", vbExclamation + vbOKOnly
      txtLastname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtFirstname.Text) = "" Then
      MsgBox "Firstname is blank.", vbExclamation + vbOKOnly
      txtFirstname.SetFocus
      Exit Sub
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
      MsgBox "Middlename is blank.", vbExclamation + vbOKOnly
      txtMiddleName.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbGender.Text) = "" Or IsNull(tdbGender.SelectedItem) Or tdbGender.ApproxCount = 0 Then
      MsgBox "Please select a gender.", vbExclamation + vbOKOnly
      tdbGender.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbCivilStatus.Text) = "" Or IsNull(tdbCivilStatus.SelectedItem) Or tdbCivilStatus.ApproxCount = 0 Then
      MsgBox "Please select a civil status.", vbExclamation + vbOKOnly
      tdbCivilStatus.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(tdbBirthdate.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      tdbBirthdate.SetFocus
      Exit Sub
    End If
    
    If Not IsDate(txtDateHired.Text) Then
      MsgBox "Please enter a valid date.", vbExclamation + vbOKOnly
      txtDateHired.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbPayFrequency.Text) = "" And IsNull(tdbPayFrequency.SelectedItem) Or tdbPayFrequency.ApproxCount = 0 Then
        MsgBox "Please assign a payroll frequency.", vbExclamation + vbOKOnly
        tdbPayFrequency.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbProvince.Text) = "" Or IsNull(tdbProvince.SelectedItem) Or tdbProvince.ApproxCount = 0 Then
        mProvince = ""
    Else
        mProvince = tdbProvince.BoundText
    End If
    
    If Trim(tdbMunicipal.Text) = "" Or IsNull(tdbMunicipal.SelectedItem) Or tdbMunicipal.ApproxCount = 0 Then
        mMunicipal = ""
    Else
        mMunicipal = tdbMunicipal.BoundText
    End If
    
    If Trim(tdbBarangay.Text) = "" Or IsNull(tdbBarangay.SelectedItem) Or tdbBarangay.ApproxCount = 0 Then
        mBarangay = ""
    Else
        mBarangay = tdbBarangay.BoundText
    End If

    If Trim(tdbBank.Text) = "" Or IsNull(tdbBank.SelectedItem) Or tdbBank.ApproxCount = 0 Then
      mBank = "Null"
    Else
      mBank = tdbBank.BoundText
    End If
    
    If Trim(tdbBranch.Text) = "" Or IsNull(tdbBranch.SelectedItem) Or tdbBranch.ApproxCount = 0 Then
      MsgBox "Please select a branch.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbBranch.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbDivision.Text) = "" Or IsNull(tdbDivision.SelectedItem) Or tdbDivision.ApproxCount = 0 Then
      MsgBox "Please select a division.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbDivision.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbCostCenter.Text) = "" Or IsNull(tdbCostCenter.SelectedItem) Or tdbDivision.ApproxCount = 0 Then
      MsgBox "Please select a cost center.", vbExclamation + vbOKOnly
      tabEmployeeInfo.CurrTab = 0
      tdbCostCenter.SetFocus
      Exit Sub
    End If
    
    If Trim(tdbRateType.Text) = "" Or IsNull(tdbRateType.SelectedItem) Or tdbRateType.ApproxCount = 0 Then
        MsgBox "Please select a rate type.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbRateType.SetFocus
        Exit Sub
    End If
    
    If Trim(tdbSection.Text) <> "" And Not IsNull(tdbSection.SelectedItem) And tdbSection.ApproxCount > 0 Then
        mSection = tdbSection.BoundText
    Else
        mSection = "Null"
    End If
    
    If Trim(tdbJob.Text) = "" Or IsNull(tdbJob.SelectedItem) Or tdbJob.ApproxCount = 0 Then
        MsgBox "Please select a job description.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbJob.SetFocus
        Exit Sub
    End If
        
    If Trim(tdbEmpStat.Text) = "" Or IsNull(tdbEmpStat.SelectedItem) Or tdbEmpStat.ApproxCount = 0 Then
        MsgBox "Please select an employment status.", vbExclamation + vbOKOnly
        tabEmployeeInfo.CurrTab = 0
        tdbEmpStat.SetFocus
        Exit Sub
    End If
    
    If optSSSAuto.Value = True Then
        mSSSAuto = 1
        mSSSAmt = 0
        mSSSEr = 0
        mSSSEc = 0
    Else
        mSSSAuto = 0
        mSSSEr = 0
        mSSSEc = 0
        mSSSAmt = Format(txtSSSAmt.Text, "###0.00")
        mSSSEr = Format(txtSssEr.Text, "###0.00")
        mSSSEc = Format(txtSssEc.Text, "###0.00")
    End If
    
    If optPhilHAuto.Value = True Then
        mPHILHAuto = 1
        mPHILHAmt = 0
        mPHILHEr = 0
    Else
        mPHILHAuto = 0
        mPHILHEr = 0
        mPHILHAmt = Format(txtPhilHAmt.Text, "###0.00")
        mPHILHEr = Format(txtPhilEr.Text, "###0.00")
    End If

    If optHDMFAuto.Value = True Then
        mHDMFAuto = 1
    Else
        mHDMFAuto = 0
        mHDMFAmt = Format(txtHdmfAmt.Text, "###0.00")
        mHdmfEr = Format(txtHDMFEr.Text, "###0.00")
    End If

    If optTaxAuto.Value = True Then
        mTAXAuto = 1
        mTAXAmt = 0
        If Trim(tdbWT.Text) = "" Or IsNull(tdbWT.SelectedItem) Or tdbWT.ApproxCount = 0 Then
            mWT = ""
        Else
            mWT = tdbWT.BoundText
        End If
    Else
        mTAXAuto = 0
        mTAXAmt = Format(txtTaxAmt.Text, "###0.00")
        mWT = ""
    End If
    
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    
    If dlgBrowsePic.FileName <> "" Then
    
      ConMain.Execute "delete from emppics where employeecode = '" & txtEmpNo.Text & "'"
      
      Set mPhoto = New ADODB.Stream
      mPhoto.Type = adTypeBinary
      mPhoto.Open
      mPhoto.LoadFromFile (dlgBrowsePic.FileName)
      
      NetOpen rsEmpPics, "select * from emppics limit 0"
      
      With rsEmpPics
        .AddNew
        .Fields("employeecode") = txtEmpNo.Text
        .Fields("images") = mPhoto.Read
        .Fields("filename") = txtBioMetId.Text & "." & GetFileExt(dlgBrowsePic.FileTitle)
        .Update
        mPicName = .Fields("filename")
      End With
      
      Set rsEmpPics = Nothing
    End If
    mEmpno = rsEmployee!employeecode
    ConMain.Execute "update employee set biometid = '" & txtBioMetId.Text & "', lastname = '" & UCase(txtLastname.Text) & "', firstname = '" & UCase(txtFirstname.Text) & "', middlename = '" & UCase(txtMiddleName.Text) & "', " & _
                          "gender = '" & tdbGender.Text & "', civilstatus = '" & tdbCivilStatus.Text & "', birthdate = '" & Format(tdbBirthdate.Text, "YYYY-MM-DD") & "', datehired = '" & Format(txtDateHired.Text, "YYYY-MM-DD") & "', houseno = '" & txtHouseNo.Text & "', street = '" & txtStreet.Text & "', " & _
                          "provcode = '" & mProvince & "', muncode = '" & mMunicipal & "', brgycode = '" & mBarangay & "', branchcode = '" & tdbBranch.BoundText & "', divisioncode = '" & tdbDivision.BoundText & "', costcentercode = '" & tdbCostCenter.BoundText & "'," & _
                          "telno = '" & txtTelno.Text & "', mobileno = '" & txtMobileno.Text & "', email = '" & txtEmail.Text & "', " & _
                          "emrgncyname = '" & txtEmrgncyName.Text & "', emrgncyno = '" & txtEmrgncyNo.Text & "', emrgncyemail = '" & txtEmrgncyEmail.Text & "',payfreqcode = '" & tdbPayFrequency.BoundText & "', filename = '" & mPicName & "', " & _
                          "wtcode = '" & mWT & "',jobtitlecode = '" & tdbJob.BoundText & "',empstatcode = '" & tdbEmpStat.BoundText & "',ratetypecode = '" & tdbRateType.BoundText & "',monthly_rate = " & Format(txtMonthly_Rate.Text, "##0.00") & ",daily_rate = " & Format(txtDaily_Rate.Text, "##0.0000000") & ",hourly_rate = " & Format(txtHourly_Rate.Text, "##0.0000000") & "," & _
                          "sssno = '" & txtSSSno.Text & "', philhno = '" & txtPhilHNo.Text & "',tinno = '" & txtTinno.Text & "',hdmfno = '" & txtHDMFNo.Text & "', bankacctno = '" & txtBankAcctNo.Text & "',sectioncode = " & mSection & ", " & _
                          "ssser = " & mSSSEr & ",sssamt = " & mSSSAmt & ",sssec = " & mSSSEc & ",philhamt = " & mPHILHAmt & ",philher = " & mPHILHEr & ",taxamt = " & mTAXAmt & ",hdmfamt = " & mHDMFAmt & ",hdmfer = " & mHdmfEr & ", " & _
                          "sssauto = " & mSSSAuto & ",philhauto = " & mPHILHAuto & ",taxauto = " & mTAXAuto & ",hdmfauto = " & mHDMFAuto & ", " & _
                          "saltobank = '" & IIf(chkSalToBank.Value <> 0, "Y", "N") & "',bankcode = " & mBank & "," & _
                          "regular = '" & IIf(ChkRegular.Value <> 0, "Y", "N") & "',isactive = '" & IIf(chkIsActive.Value <> 0, "Y", "N") & "',logbased = '" & IIf(chkLogBased.Value <> 0, "Y", "N") & "',mealallow = " & Format(txtMealAllow.Text, "##0.00") & ",fixedearnings = " & Format(txtFixedEarnings.Text, "##0.00") & " " & _
                          "where employeecode = " & rsEmployee!employeecode & ""
    
    ConMain.Execute "update payroll set saltobank = '" & IIf(chkSalToBank.Value <> 0, "Y", "N") & "' where employeecode = " & rsEmployee!employeecode & " and fnlz = 'N'"
    ConMain.Execute "delete from empshift where employeecode = '" & txtEmpNo.Text & "'"
    rsEmpShiftTmp.MoveFirst
    Do While Not rsEmpShiftTmp.EOF
      ConMain.Execute "insert into empshift(employeecode,dayno,day,shiftcode) values " & _
                            "('" & txtEmpNo.Text & "','" & rsEmpShiftTmp!dayno & "','" & rsEmpShiftTmp!Day & "', " & _
                            "'" & rsEmpShiftTmp!shiftcode & "')"
      rsEmpShiftTmp.MoveNext
    Loop
    
    ConMain.CommitTrans
    
    rsEmployee.Requery
    rsEmployee.Find "employeecode = " & mEmpno & ""
    
    dlgBrowsePic.FileName = ""
    
    cmdEmployee_Click 3
    
  End If
  
End Sub

Private Sub Cancel_Clicked()

  If cmdEmployee(0).Caption = "&Save" And cmdEmployee(0).Enabled = True Then
    tabEmployee.CurrTab = 0
  End If
  
  If rsEmployee.RecordCount > 0 Then
    Lock_Button "TTTFTT", cmdEmployee, 5
  Else
    Lock_Button "TFFFTT", cmdEmployee, 5
  End If
  
  Lock_Frame "FFFF", fra, 3
  
  Lock_Tab "TT", tabEmployee, 1

  cmdEmployee(0).Caption = "&New"
  cmdEmployee(1).Caption = "&Edit"

  fraSearch.Enabled = True
  
  tdgShift.Columns("shiftcode").Locked = True
  
  tabEmployee_Switch 0, tabEmployee.CurrTab, False
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fraButton
      .Top = TitleBar.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
    
    With tabEmployee
      .Top = fraButton.Height + TitleBar.Height
      .Left = 150
      .Width = Me.ScaleWidth - 300
      .Height = Me.ScaleHeight - .Top
    End With
    
    With fraSearch
      .Top = 0
      .Left = 150
      .Width = tabEmployee.Width - 300
    End With
    
    With tdgEmployee
      .Top = fraSearch.Top + fraSearch.Height
      .Left = 150
      .Width = tabEmployee.Width - 300
      .Height = tabEmployee.Height - (.Top + 500)
    End With
    
End Sub

Private Sub imgPhoto_Click()
    
    dlgBrowsePic.FileName = "*.jpg;*.png;*.gif"
    dlgBrowsePic.Flags = cdlOFNFileMustExist
    dlgBrowsePic.DialogTitle = "Browse Picture"
    dlgBrowsePic.ShowOpen
    If dlgBrowsePic.FileName = "*.jpg;*.png;*.gif" Then Exit Sub
    imgPhoto.Picture = LoadPicture(dlgBrowsePic.FileName)

End Sub

Private Sub optHDMFAuto_Click()
    If optHDMFAuto.Value = True Then
        txtHdmfAmt.Enabled = False
        txtHDMFEr.Enabled = False
    End If
End Sub

Private Sub optHDMFFixed_Click()
    If optHDMFFixed.Value = True Then
        txtHdmfAmt.Enabled = True
        txtHDMFEr.Enabled = True
    End If
End Sub

Private Sub optPhilHAuto_Click()
    If optPhilHAuto.Value = True Then
        txtPhilHAmt.Enabled = False
        txtPhilEr.Enabled = False
    End If
End Sub

Private Sub optPhilHFixed_Click()
    If optPhilHFixed.Value = True Then
        txtPhilHAmt.Enabled = True
        txtPhilEr.Enabled = True
    End If
End Sub

Private Sub optSSSAuto_Click()
    If optSSSAuto.Value = True Then
        txtSSSAmt.Enabled = False
        txtSssEr.Enabled = False
        txtSssEc.Enabled = False
    End If
End Sub

Private Sub optSSSFixed_Click()
    If optSSSFixed.Value = True Then
        txtSSSAmt.Enabled = True
        txtSssEr.Enabled = True
        txtSssEc.Enabled = True
    End If
End Sub

Private Sub optTaxAuto_Click()
    If optTaxAuto.Value = True Then
        tdbWT.Enabled = True
        txtTaxAmt.Enabled = False
    End If
End Sub

Private Sub optTaxFixed_Click()
    If optTaxFixed.Value = True Then
        tdbWT.Enabled = False
        txtTaxAmt.Enabled = True
    End If
End Sub

Private Sub tabEmployee_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
  
  If NewTab = 1 Then
    With rsEmployee
      If .RecordCount > 0 Then
        txtEmpNo.Text = !dummycode & ""
        txtBioMetId.Text = !biometid
        txtLastname.Text = !lastname
        txtFirstname.Text = !firstname
        txtMiddleName.Text = !middlename
        tdbGender.Text = !gender
        tdbCivilStatus.Text = !CivilStatus
        tdbBirthdate.Text = Format(!birthdate, "MM/DD/YYYY")
        txtDateHired.Text = IIf(IsNull(!datehired), "", Format(!datehired, "MM/DD/YYYY"))
        tdbPayFrequency.BoundText = !payfreqcode & ""
        If Trim(!FileName) <> "" Then
            If Dir(mEmpPicPath & "\" & !FileName) <> "" Then
                imgPhoto.Picture = LoadPicture(mEmpPicPath & "\" & !FileName) 'load image
            End If
            mPicName = !FileName
        Else
            Set imgPhoto = Nothing
            mPicName = ""
        End If
        tabEmployeeInfo_Switch 0, tabEmployeeInfo.CurrTab, False
      Else
        If cmdEmployee(0).Caption = "&New" And cmdEmployee(1).Caption = "&Edit" Then
          MsgBox "No record to view.", vbExclamation + vbOKOnly
          Cancel = True
          Exit Sub
        End If

      End If
        tabEmployeeInfo.CurrTab = 0
        'Load_Address
        
    End With
  End If
  
End Sub

Private Sub tabEmployeeInfo_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
  
    Select Case NewTab
      
        Case 0:
                If cmdEmployee(0).Caption = "&New" And cmdEmployee(1).Caption = "&Edit" Then
                    Load_Address
                End If
        Case 2:
                If cmdEmployee(0).Caption = "&New" And cmdEmployee(1).Caption = "&Edit" Then
                    Load_Deductions
                End If
        Case 1:
                If cmdEmployee(0).Caption = "&New" And cmdEmployee(1).Caption = "&Edit" Then
                    Load_EmpShift
                End If
  End Select
  
End Sub

Private Sub tdbBank_GotFocus()
  tdbBank.Tag = tdbBank.BoundText
  bind_tdb ConMain, tdbBank, "select bankcode,bankname from bank order by bankname", "bankname", "bankcode"
  tdbBank.BoundText = tdbBank.Tag
End Sub

Private Sub tdbBirthdate_Change()
  If IsDate(tdbBirthdate.Text) Then
    txtAge.Text = CInt(Format(Now, "YYYY")) - CInt(Format(tdbBirthdate.Text, "YYYY"))
  End If
End Sub

Private Sub tdbBranch_ItemChange()
  bind_tdb ConMain, tdbDivision, "select divisioncode,division from division " & _
            "where branchcode = '" & tdbBranch.BoundText & "' order by division", "division", "divisioncode"
  If tdbDivision.ApproxCount > 0 Then
    tdbDivision.BoundText = tdbDivision.Columns(0).Text
  End If
End Sub

Private Sub tdbBranch_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbBranch, tdbBranch.RowSource, tdbBranch.Text
    tdbBranch_ItemChange
  End If
End Sub

Private Sub tdbcostcenter_ItemChange()
    bind_tdb ConMain, tdbSection, "select sectioncode, sectionname from section " & _
                    "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and costcentercode = '" & tdbCostCenter.BoundText & "' order by sectionname", "sectionname", "sectioncode"
    If cmdEmployee(0).Caption = "&New" And cmdEmployee(1).Caption = "&Edit" Then
        If tdbSection.ApproxCount > 0 Then
            tdbSection.BoundText = tdbSection.Columns(0).Text
        End If
    Else
        tdbSection.BoundText = ""
    End If
End Sub

Private Sub tdbcostcenter_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbCostCenter, tdbCostCenter.RowSource, tdbCostCenter.Text
    tdbcostcenter_ItemChange
  End If
End Sub

Private Sub tdbDivision_Itemchange()
  bind_tdb ConMain, tdbCostCenter, "select costcentercode,costcenter from costcenter " & _
            "where branchcode = '" & tdbBranch.BoundText & "' order by costcenter", "costcenter", "costcentercode"
    
End Sub

Private Sub tdbDivision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbDivision, tdbDivision.RowSource, tdbDivision.Text
    tdbDivision_Itemchange
  End If
End Sub

Private Sub tdbEmpStat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbEmpStat, tdbEmpStat.RowSource, tdbEmpStat.Text
    End If
End Sub

Private Sub tdbGender_ItemChange()
    If tdbGender.BoundText = "1" Then
        chkSalToBank.Caption = "Salary will be deposited to his bank account."
        chkLogBased.Caption = "Computation of wages will be based on his time logs."
    Else
        chkSalToBank.Caption = "Salary will be deposited to her bank account."
        chkLogBased.Caption = "Computation of wages will be based on her time logs."
    End If
End Sub

Private Sub tdbGender_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbGender, tdbGender.RowSource, tdbGender.Text
    tdbGender_ItemChange
  End If
End Sub

Private Sub tdbCivilStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbCivilStatus, tdbCivilStatus.RowSource, tdbCivilStatus.Text
  End If
End Sub

Private Sub tdbSection_keyrpess(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbSection, tdbSection.RowSource, tdbSection.Text
    End If
End Sub

Private Sub tdbJob_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbJob, tdbJob.RowSource, tdbJob.Text
    End If
End Sub

Private Sub tdbBank_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      SendKeys "{TAB}"
    Else
      SearchList KeyAscii, tdbBank, tdbBank.RowSource, tdbBank.Text
    End If
End Sub

Private Sub tdbMunicipal_ItemChange()
  bind_tdb ConMain, tdbBarangay, "select brgycode, brgyname from barangay where " & _
            "provcode = '" & tdbProvince.BoundText & "' and muncode = '" & tdbMunicipal.BoundText & "' order by brgyname", "brgyname", "brgycode"
  If tdbBarangay.ApproxCount > 0 Then
    tdbBarangay.BoundText = tdbBarangay.Columns(0).Text
  End If
End Sub

Private Sub tdbBarangay_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbBarangay, tdbBarangay.RowSource, tdbBarangay.Text
  End If
End Sub

Private Sub tdbMunicipal_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbMunicipal, tdbMunicipal.RowSource, tdbMunicipal.Text
    tdbMunicipal_ItemChange
  End If
End Sub

Private Sub tdbProvince_ItemChange()
  bind_tdb ConMain, tdbMunicipal, "select muncode,munname from municipal where provcode = '" & tdbProvince.BoundText & "' order by munname", "munname", "muncode"
  If tdbMunicipal.ApproxCount > 0 Then
    tdbMunicipal.BoundText = tdbMunicipal.Columns(0).Text
  End If
End Sub

Private Sub tdbProvince_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbProvince, tdbProvince.RowSource, tdbProvince.Text
    tdbProvince_ItemChange
  End If
End Sub

Private Sub tdbRateType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbRateType, tdbRateType.RowSource, tdbRateType.Text
    End If
End Sub

Private Sub tdbSort_ItemChange()
    rsEmployee.Sort = tdbSort.BoundText
    mSort = tdbSort.BoundText
End Sub

Private Sub tdbSort_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        SearchList KeyAscii, tdbSort, tdbSort.RowSource, tdbSort.Text
        tdbSort_ItemChange
    End If
End Sub

Private Sub tdbWT_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbWT, tdbWT.RowSource, tdbWT.Text
    End If
End Sub

Private Sub tddShift_DropDownOpen()
    Bind_tdd ConMain, tddShift, "select shiftcode, concat(t1in, ' ', t1out, ' ', t2in, ' ',t2out)  shiftdesc,t1in,t1out,t2in,t2out from shift", "shiftcode"
End Sub

Private Sub tddShift_RowChange()
  With tdgShift
    txtShiftcode.Text = tddShift.Columns("shiftcode").Text
    .Columns("t1in").Text = tddShift.Columns("t1in").Text
    .Columns("t1out").Text = tddShift.Columns("t1out").Text
    .Columns("t2in").Text = tddShift.Columns("t2in").Text
    .Columns("t2out").Text = tddShift.Columns("t2out").Text
  End With
End Sub

Private Sub tdgShift_KeyDown(KeyCode As Integer, Shift As Integer)
    
  If cmdEmployee(0).Caption <> "&New" Or cmdEmployee(1).Caption <> "&Edit" Then
    If KeyCode = 46 Then
      With tdgShift
        .Columns("shiftcode").Text = ""
        .Columns("t1in").Text = ""
        .Columns("t1out").Text = ""
        .Columns("t2in").Text = ""
        .Columns("t2out").Text = ""
      End With
    End If
  End If
End Sub

Private Sub txtFixedEarnings_GotFocus()
    With txtFixedEarnings
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
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
  Else
    SearchRecord KeyAscii, txtSearch, rsEmployee, txtSearch.Text, mSort
  End If
End Sub

Private Sub ClearText()

      txtEmpNo.Text = ""
      txtBioMetId.Text = ""
      txtLastname.Text = ""
      txtFirstname.Text = ""
      txtMiddleName.Text = ""
      tdbGender.Text = ""
      tdbCivilStatus.Text = ""
      tdbBirthdate.Text = Format(Now, "MM/DD/YYYY")
      txtDateHired.Text = Format(Now, "MM/DD/YYYY")
      tdbPayFrequency.Text = ""
      txtHouseNo.Text = ""
      txtStreet.Text = ""
      tdbProvince.BoundText = ""
      tdbMunicipal.BoundText = ""
      tdbBarangay.BoundText = ""
      tdbBranch.BoundText = ""
      tdbDivision.BoundText = ""
      tdbCostCenter.BoundText = ""
      txtTelno.Text = ""
      txtMobileno.Text = ""
      txtEmail.Text = ""
      txtEmrgncyName.Text = ""
      txtEmrgncyNo.Text = ""
      txtEmrgncyEmail.Text = ""
      tdbWT.BoundText = ""
      tdbSection.BoundText = ""
      tdbJob.BoundText = ""
      tdbEmpStat.BoundText = ""
      tdbRateType.BoundText = ""
      tdbBank.BoundText = ""
      txtMonthly_Rate.Text = "0.00"
      txtDaily_Rate.Text = "0.0000000"
      txtHourly_Rate.Text = "0.0000000"
      txtMealAllow.Text = "0.00"
      txtFixedEarnings.Text = "0.00"
      txtSSSno.Text = ""
      txtPhilHNo.Text = ""
      txtTinno.Text = ""
      txtBankAcctNo.Text = ""
      imgPhoto.Picture = Nothing

End Sub

Private Sub Load_Address()

  With rsEmployee
            
        bind_tdb ConMain, tdbPayFrequency, "select payfreqcode,payfreqname from payfrequency order by payfreqname", "payfreqname", "payfreqcode"
        bind_tdb ConMain, tdbProvince, "select provcode, provname from province order by provname", "provname", "provcode"
        bind_tdb ConMain, tdbBranch, "select branchcode, branch from branch order by branch", "branch", "branchcode"
      
      If .RecordCount > 0 Then
        
        tdbPayFrequency.BoundText = !payfreqcode & ""
        txtHouseNo.Text = !houseno
        txtStreet.Text = !street
        tdbProvince.BoundText = !provcode & ""
        DoEvents
        bind_tdb ConMain, tdbMunicipal, "select muncode, munname from municipal where provcode = '" & !provcode & "' order by munname", "munname", "muncode"
        DoEvents
        tdbMunicipal.BoundText = !muncode & ""
        DoEvents
        bind_tdb ConMain, tdbBarangay, "select brgycode, brgyname from barangay " & _
                "where provcode = '" & !provcode & "' and muncode = '" & !muncode & "' order by brgyname", "brgyname", "brgycode"
        
        DoEvents
        tdbBarangay.BoundText = !brgycode & ""
        DoEvents
        tdbBranch.BoundText = !branchcode
        DoEvents
        bind_tdb ConMain, tdbDivision, "select divisioncode, division from division where branchcode = '" & !branchcode & "' order by division", "division", "divisioncode"
        DoEvents
        tdbDivision.BoundText = !divisioncode
        DoEvents
        bind_tdb ConMain, tdbCostCenter, "select costcentercode, costcenter from costcenter  order by costcenter", "costcenter", "costcentercode"
        DoEvents
        tdbCostCenter.BoundText = !costcentercode & ""
        DoEvents
        bind_tdb ConMain, tdbSection, "select sectioncode, sectionname from section " & _
                "where branchcode = '" & !branchcode & "' and divisioncode = '" & !divisioncode & "' and costcentercode = '" & !costcentercode & "' order by sectionname", "sectionname", "sectioncode"
        DoEvents
        tdbSection.BoundText = !sectioncode & ""
        DoEvents
        
        tdbJob.BoundText = !JobTitlecode
        tdbEmpStat.BoundText = !empstatcode
        tdbRateType.BoundText = !ratetypecode
        tdbBank.BoundText = !bankcode & ""
        txtMonthly_Rate.Text = IIf(IsNull(!Monthly_Rate), "0.00", Format(!Monthly_Rate, "#,##0.00"))
        txtDaily_Rate.Text = IIf(IsNull(!Daily_Rate), "0.0000000", Format(!Daily_Rate, "#,##0.0000000"))
        txtHourly_Rate.Text = IIf(IsNull(!Hourly_Rate), "0.0000000", Format(!Hourly_Rate, "#,##0.0000000"))
        txtMealAllow.Text = IIf(IsNull(!MealAllow), "0.00", Format(!MealAllow, "#,##0.00"))
        txtFixedEarnings.Text = IIf(IsNull(!FixedEarnings), "0.00", Format(!FixedEarnings, "#,##0.00"))
        txtBankAcctNo.Text = Trim(!bankacctno) & ""
        
        chkSalToBank.Value = IIf(!saltobank = "Y", 1, 0)
        ChkRegular.Value = IIf(!regular = "Y", 1, 0)
        chkIsActive.Value = IIf(!isactive = "Y", 1, 0)
        
        txtTelno.Text = !telno & ""
        txtMobileno.Text = !mobileno & ""
        txtEmail.Text = !email & ""
        txtEmrgncyName.Text = !EmrgncyName & ""
        txtEmrgncyNo.Text = !EmrgncyNo & ""
        txtEmrgncyEmail.Text = !EmrgncyEmail & ""
        
      Else
        
        bind_tdb ConMain, tdbMunicipal, "select muncode,munname from municipal where provcode = '" & tdbProvince.BoundText & "' order by munname", "munname", "muncode"
        bind_tdb ConMain, tdbBarangay, "select brgycode, brgyname from barangay " & _
                "where provcode = '" & tdbProvince.BoundText & "' and muncode = '" & tdbMunicipal.BoundText & "' order by brgyname", "brgyname", "brgycode"
        bind_tdb ConMain, tdbDivision, "select divisioncode, division from division where branchcode = '" & tdbBranch.BoundText & "' order by division", "division", "divisioncode"
        bind_tdb ConMain, tdbCostCenter, "select costcentercode, costcenter from costcenter where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' order by costcenter", "costcenter", "costcentercode"
        bind_tdb ConMain, tdbSection, "select sectioncode, sectionname from section " & _
                "where branchcode = '" & tdbBranch.BoundText & "' and divisioncode = '" & tdbDivision.BoundText & "' and costcentercode = '" & tdbCostCenter.BoundText & "' order by sectionname", "sectionname", "sectioncode"
        
      End If 'L3
          
  End With

End Sub

Private Sub Load_Deductions()

    With rsEmployee
    
        tdbWT.BoundText = !wtcode & ""
        txtSSSno.Text = !sssno & ""
        txtPhilHNo.Text = !philhno & ""
        txtTinno.Text = !tinno & ""
        txtHDMFNo.Text = !hdmfno & ""
        txtSSSAmt.Text = Format(!sssamt, "#,##0.00")
        txtSssEr.Text = Format(!SssEr, "#,##0.00")
        txtSssEc.Text = Format(!SssEc, "#,##0.00")
        txtPhilHAmt.Text = Format(!PhilHAmt, "#,##0.00")
        txtPhilEr.Text = Format(!philher, "#,##0.00")
        txtHdmfAmt.Text = Format(!HdmfAmt, "#,##0.00")
        txtHDMFEr.Text = Format(!HDMFEr, "#,##0.00")
        txtTaxAmt.Text = Format(!taxamt, "#,##0.00")
        
        If !sssauto <> 0 Then
            optSSSAuto.Value = True
        Else
            optSSSFixed.Value = True
        End If
        If !philhauto <> 0 Then
            optPhilHAuto.Value = True
        Else
            optPhilHFixed.Value = True
        End If
        If !hdmfauto <> 0 Then
            optHDMFAuto.Value = True
        Else
            optHDMFFixed.Value = True
        End If
        If !taxauto <> 0 Then
            optTaxAuto.Value = True
        Else
            optTaxFixed.Value = True
        End If
    End With

End Sub

Private Sub Load_EmpShift()
  
    With rsEmployee
        If .RecordCount > 0 Then
            chkLogBased.Value = IIf(!logbased = "Y", 1, 0)
            NetOpen rsEmpShift, "select x1.employeecode,x1.dayno,x1.day,x2.* from empshift x1 left outer join shift x2 on x1.shiftcode = x2.shiftcode where x1.employeecode = '" & !employeecode & "' order by x1.dayno"
            If rsEmpShift.RecordCount > 0 Then
                rsEmpShift.MoveFirst
                rsEmpShiftTmp.MoveFirst
                Do While Not rsEmpShift.EOF
                    rsEmpShiftTmp.Fields("shiftcode") = rsEmpShift!shiftcode & ""
                    rsEmpShiftTmp.Fields("t1in") = rsEmpShift!t1in & ""
                    rsEmpShiftTmp.Fields("t1out") = rsEmpShift!t1out & ""
                    rsEmpShiftTmp.Fields("t2in") = rsEmpShift!t2in & ""
                    rsEmpShiftTmp.Fields("t2out") = rsEmpShift!t2out & ""
                    rsEmpShiftTmp.MoveNext
                    rsEmpShift.MoveNext
                Loop
            End If
        End If
    End With

End Sub

Private Sub CreateShiftTmp(ByRef rs As ADODB.Recordset)

  Set rs = Nothing
  Set rs = New ADODB.Recordset
  With rs
  End With
  
End Sub

Private Sub CreateEmpShiftTmp()
  
  Dim i             As Integer
  
  Set rsEmpShiftTmp = Nothing
  Set rsEmpShiftTmp = New ADODB.Recordset
  
  With rsEmpShiftTmp
    .Fields.Append "dayno", adVarChar, 1
    .Fields.Append "day", adVarChar, 20
    .Fields.Append "shiftcode", adVarChar, 7
    .Fields.Append "t1in", adVarChar, 5
    .Fields.Append "t1out", adVarChar, 5
    .Fields.Append "t2in", adVarChar, 5
    .Fields.Append "t2out", adVarChar, 5
    .Open
  
    For i = 1 To 7
      .AddNew
      .Fields("Dayno") = i
      Select Case i
        Case 1: .Fields("day") = "Sunday"
        Case 2: .Fields("day") = "Monday"
        Case 3: .Fields("day") = "Tuesday"
        Case 4: .Fields("day") = "Wednesday"
        Case 5: .Fields("day") = "Thursday"
        Case 6: .Fields("day") = "Friday"
        Case 7: .Fields("day") = "Saturday"
      End Select
      .Update
    Next
  
  End With
  
  
  Set tdgShift.DataSource = rsEmpShiftTmp
  
  
End Sub

Private Sub txtShiftcode_LostFocus()
  tdgShift.SetFocus
End Sub

Private Sub txtshiftcode_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchRecord KeyAscii, txtShiftcode, tddShift.DataSource, txtShiftcode.Text, "shiftcode"
    tddShift_RowChange
  End If
End Sub
