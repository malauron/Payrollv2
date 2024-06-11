VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAdEmployeePerformanceEvaluation2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Employee Performance Evaluation"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -30
      TabIndex        =   28
      Top             =   7230
      Width           =   8775
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   60
         TabIndex        =   22
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
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
         Image           =   "frmAdEmployeePerformanceEvaluation2.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   1800
         TabIndex        =   21
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
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
         Image           =   "frmAdEmployeePerformanceEvaluation2.frx":0CDA
         cBack           =   14737632
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   7320
      Left            =   30
      TabIndex        =   23
      Top             =   -75
      Width           =   8595
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SCORE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3345
         Index           =   4
         Left            =   7335
         TabIndex        =   58
         Top             =   3090
         Width           =   900
         Begin VB.Label lblTtlScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   75
            TabIndex        =   68
            Top             =   2940
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   6
            Left            =   75
            TabIndex        =   65
            Top             =   2580
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   64
            Top             =   2220
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   63
            Top             =   1860
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   3
            Left            =   75
            TabIndex        =   62
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   61
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   60
            Top             =   780
            Width           =   735
         End
         Begin VB.Label lblScore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   59
            Top             =   420
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RATING"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3345
         Index           =   3
         Left            =   6300
         TabIndex        =   57
         Top             =   3090
         Width           =   1035
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   13
            Top             =   420
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":19B4
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":19D4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1A40
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1A5E
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":1AA8
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   14
            Top             =   780
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":1AD0
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":1AF0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1B5C
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1B7A
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":1BC4
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   15
            Top             =   1140
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":1BEC
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":1C0C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1C78
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1C96
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":1CE0
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   3
            Left            =   75
            TabIndex        =   16
            Top             =   1500
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":1D08
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":1D28
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1D94
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1DB2
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":1DFC
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   17
            Top             =   1860
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":1E24
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":1E44
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1EB0
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1ECE
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":1F18
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   18
            Top             =   2220
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":1F40
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":1F60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":1FCC
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":1FEA
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":2034
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRating 
            Height          =   315
            Index           =   6
            Left            =   75
            TabIndex        =   19
            Top             =   2580
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            Calculator      =   "frmAdEmployeePerformanceEvaluation2.frx":205C
            Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":207C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":20E8
            Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":2106
            Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":2150
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4210752
            Format          =   "##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   5
            MinValue        =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   1
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MULTIPLIER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3345
         Index           =   2
         Left            =   4860
         TabIndex        =   49
         Top             =   3090
         Width           =   1440
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   6
            Left            =   75
            TabIndex        =   56
            Top             =   2580
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   55
            Top             =   2220
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   54
            Top             =   1860
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   3
            Left            =   75
            TabIndex        =   53
            Top             =   1500
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   2
            Left            =   75
            TabIndex        =   52
            Top             =   1140
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   51
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   50
            Top             =   420
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "WEIGHT SCORE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3345
         Index           =   1
         Left            =   3135
         TabIndex        =   41
         Top             =   3090
         Width           =   1725
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   100%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   18
            Left            =   75
            TabIndex        =   67
            Top             =   2940
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   17
            Left            =   75
            TabIndex        =   48
            Top             =   420
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   15
            Left            =   75
            TabIndex        =   47
            Top             =   780
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   14
            Left            =   75
            TabIndex        =   46
            Top             =   1140
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   13
            Left            =   75
            TabIndex        =   45
            Top             =   1500
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   12
            Left            =   75
            TabIndex        =   44
            Top             =   1860
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   11
            Left            =   75
            TabIndex        =   43
            Top             =   2220
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   10
            Left            =   75
            TabIndex        =   42
            Top             =   2580
            Width           =   1560
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "GENERAL CRITERIA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   3345
         Index           =   0
         Left            =   1035
         TabIndex        =   33
         Top             =   3090
         Width           =   2100
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Job Knowledge"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   16
            Left            =   75
            TabIndex        =   40
            Top             =   2580
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Self Control"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   9
            Left            =   75
            TabIndex        =   39
            Top             =   2220
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Work Habits"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   8
            Left            =   75
            TabIndex        =   38
            Top             =   1860
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Innitiative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   7
            Left            =   75
            TabIndex        =   37
            Top             =   1500
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Communication"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   6
            Left            =   75
            TabIndex        =   36
            Top             =   1140
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Attendance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   35
            Top             =   780
            Width           =   1950
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Personality"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   315
            Index           =   4
            Left            =   75
            TabIndex        =   34
            Top             =   420
            Width           =   1950
         End
      End
      Begin TDBText6Ctl.TDBText txtRemarks 
         Height          =   780
         Left            =   1050
         TabIndex        =   20
         Tag             =   "txtRegistrationRemarks"
         Top             =   6465
         Width           =   7170
         _Version        =   65536
         _ExtentX        =   12647
         _ExtentY        =   1376
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2178
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":21E4
         Key             =   "frmAdEmployeePerformanceEvaluation2.frx":2202
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
         AlignVertical   =   0
         MultiLine       =   -1
         ScrollBars      =   2
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
      Begin TDBText6Ctl.TDBText txtJobTitle 
         Height          =   300
         Left            =   2205
         TabIndex        =   2
         Top             =   570
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   529
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2246
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":22B2
         Key             =   "frmAdEmployeePerformanceEvaluation2.frx":22D0
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
      Begin TDBDate6Ctl.TDBDate txtEvalDate 
         Height          =   315
         Left            =   2205
         TabIndex        =   9
         Top             =   1965
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         Calendar        =   "frmAdEmployeePerformanceEvaluation2.frx":2314
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":241A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":2480
         Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":249E
         Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":24FC
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
      Begin TDBDate6Ctl.TDBDate txtEvalFrom 
         Height          =   315
         Left            =   2205
         TabIndex        =   10
         Top             =   2325
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         Calendar        =   "frmAdEmployeePerformanceEvaluation2.frx":2524
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":262A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":2690
         Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":26AE
         Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":270C
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
      Begin TDBDate6Ctl.TDBDate txtEvalTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   11
         Top             =   2325
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         Calendar        =   "frmAdEmployeePerformanceEvaluation2.frx":2734
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":283A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":28A0
         Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":28BE
         Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":291C
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
      Begin TDBText6Ctl.TDBText txtCostCenter 
         Height          =   300
         Left            =   2205
         TabIndex        =   6
         Top             =   1260
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   529
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2944
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":29B0
         Key             =   "frmAdEmployeePerformanceEvaluation2.frx":29CE
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
      Begin TDBText6Ctl.TDBText txtEmployee 
         Height          =   300
         Left            =   2205
         TabIndex        =   0
         Top             =   225
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   529
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2A12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":2A7E
         Key             =   "frmAdEmployeePerformanceEvaluation2.frx":2A9C
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         ToolTipText     =   "Browse for checked in guests."
         Top             =   225
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
      Begin lvButton.lvButtons_H cmdJobTitle 
         Height          =   315
         Left            =   8190
         TabIndex        =   3
         ToolTipText     =   "Browse for checked in guests."
         Top             =   570
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
      Begin lvButton.lvButtons_H cmdCostCenter 
         Height          =   315
         Left            =   8190
         TabIndex        =   7
         ToolTipText     =   "Browse for checked in guests."
         Top             =   1260
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
      Begin TrueOleDBList80.TDBCombo tdbTmp 
         Bindings        =   "frmAdEmployeePerformanceEvaluation2.frx":2AE0
         DataMember      =   "tdbJob"
         Height          =   300
         Left            =   5535
         TabIndex        =   66
         Top             =   2565
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
         AnimateWindow   =   3
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   2
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"frmAdEmployeePerformanceEvaluation2.frx":2AF1
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
      Begin TrueOleDBList80.TDBCombo tdbEvalType 
         Bindings        =   "frmAdEmployeePerformanceEvaluation2.frx":2B9B
         DataMember      =   "tdbJob"
         Height          =   300
         Left            =   2205
         TabIndex        =   12
         Top             =   2685
         Width           =   2655
         _ExtentX        =   4683
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
         AnimateWindow   =   3
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   2
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"frmAdEmployeePerformanceEvaluation2.frx":2BAC
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
      Begin TDBText6Ctl.TDBText txtDivision 
         Height          =   300
         Left            =   2205
         TabIndex        =   4
         Top             =   915
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   529
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2C56
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":2CC2
         Key             =   "frmAdEmployeePerformanceEvaluation2.frx":2CE0
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
      Begin lvButton.lvButtons_H cmdDivision 
         Height          =   315
         Left            =   8190
         TabIndex        =   5
         ToolTipText     =   "Browse for checked in guests."
         Top             =   915
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
      Begin TDBDate6Ctl.TDBDate txtDateHired 
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   1605
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   556
         Calendar        =   "frmAdEmployeePerformanceEvaluation2.frx":2D24
         Caption         =   "frmAdEmployeePerformanceEvaluation2.frx":2E2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmAdEmployeePerformanceEvaluation2.frx":2E90
         Keys            =   "frmAdEmployeePerformanceEvaluation2.frx":2EAE
         Spin            =   "frmAdEmployeePerformanceEvaluation2.frx":2F0C
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
      Begin TrueOleDBList80.TDBCombo tdbEmp 
         Bindings        =   "frmAdEmployeePerformanceEvaluation2.frx":2F34
         DataMember      =   "tdbJob"
         Height          =   300
         Left            =   5535
         TabIndex        =   71
         Top             =   2235
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
         AnimateWindow   =   3
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   2
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"frmAdEmployeePerformanceEvaluation2.frx":2F45
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   19
         Left            =   195
         TabIndex        =   70
         Top             =   1650
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   690
         TabIndex        =   69
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Evaluation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   -315
         TabIndex        =   32
         Top             =   2715
         Width           =   2445
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   690
         TabIndex        =   31
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   2370
         Width           =   345
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Period of Evaluation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   -315
         TabIndex        =   29
         Top             =   2370
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   15
         TabIndex        =   27
         Top             =   6480
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   26
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   690
         TabIndex        =   25
         Top             =   615
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Evaluation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   24
         Top             =   2010
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAdEmployeePerformanceEvaluation2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd             As Boolean

Dim mTxt                As TDBText

Private Sub cmdClose_Click()
    
    Unload Me
    
    frmAdEmployeePerformanceEvaluation.tdgEmpEval.SetFocus
    
End Sub

Private Sub cmdCostCenter_Click()
  If IsNumeric(txtDivision.Tag) Then
    bind_tdb ConMain, tdbTmp, "select CostCentercode,CostCenter from CostCenter where divisioncode = " & txtDivision.Tag & " order by CostCenter", "CostCenter", "CostCentercode"
    Set mTxt = txtCostCenter
    tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
    tdbTmp.BoundText = mTxt.Tag
    mTxt.Visible = False
    tdbTmp.Visible = True
    tdbTmp.SetFocus
    SendKeys "{F4}"
  Else
    MsgBox "Please select a division.", vbExclamation + vbOKOnly
    cmdDivision.SetFocus
  End If
End Sub

Private Sub cmdDivision_Click()
  
  bind_tdb ConMain, tdbTmp, "select divisioncode,division from division order by division", "division", "divisioncode"
  Set mTxt = txtDivision
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"

End Sub

Private Sub cmdEmployee_Click()
  
  bind_tdb ConMain, tdbEmp, "select employeecode,concat(lastname,', ',firstname,' ',middlename) employeename from employee order by employeename", "employeename", "employeecode"
  Set mTxt = txtEmployee
  tdbEmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbEmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbEmp.Visible = True
  tdbEmp.SetFocus
  SendKeys "{F4}"
  
End Sub

Private Sub cmdJobTitle_Click()
  
  bind_tdb ConMain, tdbTmp, "select jobtitlecode,jobtitle from jobtitle order by jobtitle", "jobtitle", "jobtitlecode"
  Set mTxt = txtJobTitle
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrHndlr
    
    Dim rsDate                          As ADODB.Recordset
    
    Dim mEmpEvalCode                    As Double
    Dim mTotal                          As Double
    
    Dim mTransStatus                    As Boolean
    
    Dim mCostCenter                     As String
    
    mTransStatus = False
    
    
    With frmAdEmployeePerformanceEvaluation
    
        If Not IsNumeric(txtEmployee.Tag) Then
          MsgBox "Please select an employee.", vbExclamation + vbOKOnly
          cmdEmployee.SetFocus
          Exit Sub
        End If
        
        If Not IsNumeric(txtJobTitle.Tag) Then
          MsgBox "Job title is blank.", vbExclamation + vbOKOnly
          cmdJobTitle.SetFocus
          Exit Sub
        End If
        
        If Not IsNumeric(txtDivision.Tag) Then
          MsgBox "Please division is blank.", vbExclamation + vbOKOnly
          cmdDivision.SetFocus
          Exit Sub
        End If
        
        If Not IsNumeric(txtCostCenter.Tag) Then
          mCostCenter = "Null"
        Else
          mCostCenter = txtCostCenter.Tag
        End If
        
        If Not IsDate(txtDateHired.Text) Then
          MsgBox "Invalid date format.", vbExclamation + vbOKOnly
          txtDateHired.SetFocus
          Exit Sub
        End If
    
        If Not IsDate(txtEvalDate.Text) Then
          MsgBox "Invalid date format.", vbExclamation + vbOKOnly
          txtEvalDate.SetFocus
          Exit Sub
        End If
    
        If Not IsDate(txtEvalFrom.Text) Then
          MsgBox "Invalid date format.", vbExclamation + vbOKOnly
          txtEvalFrom.SetFocus
          Exit Sub
        End If
        
        If Not IsDate(txtEvalTo.Text) Then
          MsgBox "Invalid date format.", vbExclamation + vbOKOnly
          txtEvalTo.SetFocus
          Exit Sub
        End If
        
        If Trim(tdbEvalType.Text) = "" Or IsNull(tdbEvalType.SelectedItem) Or tdbEvalType.ApproxCount <= 0 Then
          MsgBox "Pleaes select an evaluation type.", vbExclamation + vbOKOnly
          tdbEvalType.SetFocus
          Exit Sub
        End If
        
        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans
        mTransStatus = True
        
        NetOpen rsDate, "select curdate() currentdate"
        If mAdd = True Then
            mEmpEvalCode = LastCode("EmpEval" & Format(rsDate!currentdate, "YY"))
            mEmpEvalCode = CDbl(Format(rsDate!currentdate, "YY") & Format(mEmpEvalCode, "00000000"))
            ConMain.Execute "insert into empeval (empevalcode,employeecode,jobtitlecode,divisioncode,costcentercode," & _
                    "trnxdate,hireddate,evaluationdate,evalfrom," & _
                    "evalto,evaltype,personality,attendance,communication," & _
                    "innitiative,workhabits,selfcontrol,jobknowledge,remarks) values (" & _
                    "" & mEmpEvalCode & "," & txtEmployee.Tag & "," & txtJobTitle.Tag & "," & txtDivision.Tag & "," & mCostCenter & "," & _
                    "Now(),'" & Format(txtDateHired.Text, "YYYY-MM-DD") & "','" & Format(txtEvalDate.Text, "YYYY-MM-DD") & "','" & Format(txtEvalFrom.Text, "YYYY-MM-DD") & "'," & _
                    "'" & Format(txtEvalTo.Text, "YYYY-MM-DD") & "','" & tdbEvalType.BoundText & "'," & txtRating(0).Text & "," & txtRating(1).Text & "," & txtRating(2).Text & "," & _
                    "" & txtRating(3).Text & "," & txtRating(4).Text & "," & txtRating(5).Text & "," & txtRating(6).Text & ",'" & Swap(txtRemarks.Text) & "')"
        
        Else
              mEmpEvalCode = .rsEmpEval!EmpEvalcode
              ConMain.Execute "update empeval set jobtitlecode = " & txtJobTitle.Tag & ",divisioncode = " & txtDivision.Tag & ",costcentercode = " & mCostCenter & ",hireddate = '" & Format(txtDateHired.Text, "YYYY-MM-DD") & "', " & _
                    "remarks = '" & Swap(txtRemarks.Text) & "', evaluationdate = '" & Format(txtEvalDate.Text, "YYYY-MM-DD") & "',evalfrom = '" & Format(txtEvalFrom.Text, "YYYY-MM-DD") & "'," & _
                    "evalto = '" & Format(txtEvalTo.Text, "YYYY-MM-DD") & "',evaltype = '" & tdbEvalType.BoundText & "',personality = " & txtRating(0).Text & ",attendance = " & txtRating(1).Text & ",communication = " & txtRating(2).Text & "," & _
                    "innitiative = " & txtRating(3).Text & ",workhabits = " & txtRating(4).Text & ",selfcontrol = " & txtRating(5).Text & ",jobknowledge = " & txtRating(6).Text & " where empevalcode = " & mEmpEvalCode & ""
        End If
        
        ConMain.CommitTrans
        mTransStatus = False
        
        If Format(rsDate!currentdate, "MM/DD/YYYY") = .txtTrnxDate.Text Then
          .rsEmpEval.Requery
          If .rsEmpEval.RecordCount > 0 Then
              Lock_Button "TTFTTT", .cmdMenu, 5
          Else
              Lock_Button "TFFFTT", .cmdMenu, 5
          End If
          .rsEmpEval.MoveFirst
          .rsEmpEval.Find "EmpEvalCode = " & mEmpEvalCode & ""
        End If
        
        Clear_Data
    End With
    Exit Sub
ErrHndlr:
    
    MsgBox "Error Message: " & err.Description, vbCritical + vbOKOnly
    If mTransStatus Then ConMain.RollbackTrans
End Sub

Private Sub Form_Activate()
  On Error Resume Next
'  tdbParticulars.SetFocus
End Sub

Private Sub Form_Load()
      
    Dim rsTmp             As ADODB.Recordset
    
    Dim i                 As Integer
    
    CreateTmpDB rsTmp
    
    With rsTmp
      For i = 1 To 4
        .AddNew
        Select Case i
            Case 1: .Fields("code") = "Annual"
                    .Fields("description") = "Annual"
            Case 2: .Fields("code") = "Probationary"
                    .Fields("description") = "Probationary"
            Case 3: .Fields("code") = "Trainee"
                    .Fields("description") = "Trainee"
            Case 4: .Fields("code") = "Others"
                    .Fields("description") = "Others"
        End Select
        .Update
      Next
    End With
    
    With tdbEvalType
      .BoundColumn = "CODE"
      .ListField = "Description"
      .Columns(0).DataField = "CODE"
      .Columns(1).DataField = "Description"
      .RowSource = rsTmp
    End With
    
    Set rsTmp = Nothing
    
    If mAdd = False Then
      cmdEmployee.Enabled = False
'      cmdJobTitle.Enabled = False
'      cmdDivision.Enabled = False
'      cmdCostCenter.Enabled = False
      With frmAdEmployeePerformanceEvaluation
        txtEmployee.Text = .rsEmpEval!employeename
        txtEmployee.Tag = .rsEmpEval!employeecode
        NetOpen rsTmp, "select jobtitle from jobtitle where jobtitlecode = " & .rsEmpEval!jobtitlecode & ""
        If rsTmp.RecordCount > 0 Then txtJobTitle.Text = rsTmp!jobtitle & ""
        txtJobTitle.Tag = .rsEmpEval!jobtitlecode
        NetOpen rsTmp, "select division from division where divisioncode = " & .rsEmpEval!divisioncode & ""
        If rsTmp.RecordCount > 0 Then txtDivision.Text = rsTmp!Division & ""
        txtDivision.Tag = .rsEmpEval!divisioncode
        If Not IsNull(.rsEmpEval!costcentercode) Then
          NetOpen rsTmp, "select costcenter from costcenter where costcentercode = " & .rsEmpEval!costcentercode & ""
          If rsTmp.RecordCount > 0 Then txtCostCenter.Text = rsTmp!CostCenter & ""
          txtCostCenter.Tag = .rsEmpEval!costcentercode
        Else
          txtCostCenter.Text = ""
          txtCostCenter.Tag = ""
        End If
        txtDateHired.Text = Format(.rsEmpEval!hireddate, "MM/DD/YYYY")
        txtEvalDate.Text = Format(.rsEmpEval!evaluationdate, "MM/DD/YYYY")
        txtEvalFrom.Text = Format(.rsEmpEval!evalfrom, "MM/DD/YYYY")
        txtEvalTo.Text = Format(.rsEmpEval!evalto, "MM/DD/YYYY")
        tdbEvalType.BoundText = .rsEmpEval!evaltype
        
        i = 0
        
        For i = txtRating.LBound To txtRating.UBound
          Select Case i:
            Case 0: txtRating(i).Text = .rsEmpEval!personality
            Case 1: txtRating(i).Text = .rsEmpEval!attendance
            Case 2: txtRating(i).Text = .rsEmpEval!communication
            Case 3: txtRating(i).Text = .rsEmpEval!innitiative
            Case 4: txtRating(i).Text = .rsEmpEval!workhabits
            Case 5: txtRating(i).Text = .rsEmpEval!selfcontrol
            Case 6: txtRating(i).Text = .rsEmpEval!jobknowledge
          End Select
          txtRating_LostFocus i
        Next
        
        
        
        txtRemarks.Text = .rsEmpEval!remarks & ""
        
      End With
    End If
    
End Sub

Private Sub tdbEmp_KeyPress(KeyAscii As Integer)
  
  Dim rsEmpTmp          As ADODB.Recordset
  Dim rsTmp             As ADODB.Recordset
  
  If KeyAscii = 13 Then
    If Trim(tdbEmp.Text) <> "" And Not IsNull(tdbEmp.SelectedItem) And tdbEmp.ApproxCount > 0 Then
      mTxt.Tag = tdbEmp.BoundText
      mTxt.Text = tdbEmp.Text
      
      NetOpen rsEmpTmp, "select jobtitlecode,divisioncode,costcentercode,datehired from employee where employeecode = " & mTxt.Tag & ""
      If rsEmpTmp.RecordCount > 0 Then
        NetOpen rsTmp, "select jobtitle from jobtitle where jobtitlecode = " & rsEmpTmp!jobtitlecode & ""
        If rsTmp.RecordCount > 0 Then txtJobTitle.Text = rsTmp!jobtitle & ""
        txtJobTitle.Tag = rsEmpTmp!jobtitlecode
        NetOpen rsTmp, "select division from division where divisioncode = " & rsEmpTmp!divisioncode & ""
        If rsTmp.RecordCount > 0 Then txtDivision.Text = rsTmp!Division & ""
        txtDivision.Tag = rsEmpTmp!divisioncode
        
        If Not IsNull(rsEmpTmp!costcentercode) Then
          NetOpen rsTmp, "select costcenter from costcenter where costcentercode = " & rsEmpTmp!costcentercode & ""
          If rsTmp.RecordCount > 0 Then txtCostCenter.Text = rsTmp!CostCenter & ""
          txtCostCenter.Tag = rsEmpTmp!costcentercode
        Else
          txtCostCenter.Text = ""
          txtCostCenter.Tag = ""
        End If
        
        If Not IsNull(rsEmpTmp!datehired) Then
          txtDateHired.Text = Format(rsEmpTmp!datehired, "MM/DD/YYYY")
        Else
          txtDateHired.Text = ""
        End If
      End If
    Else
      mTxt.Tag = ""
      mTxt.Text = ""
    End If
    mTxt.Visible = True
    mTxt.SetFocus
    tdbEmp.Visible = False
  Else
    SearchList KeyAscii, tdbEmp, tdbEmp.RowSource, tdbEmp.Text
  End If
End Sub

Private Sub tdbEvalType_GotFocus()
  With tdbEvalType
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub tdbEvalType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  Else
    SearchList KeyAscii, tdbEvalType, tdbEvalType.RowSource, tdbEvalType.Text
  End If
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

Private Sub txtCostCenter_GotFocus()
  With txtCostCenter
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCostCenter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDateHired_GotFocus()
  With txtDateHired
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDateHired_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDivision_GotFocus()
  With txtDivision
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDivision_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEmployee_GotFocus()
  With txtEmployee
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEvalFrom_GotFocus()
  With txtEvalFrom
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEvalFrom_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtEvalTo_GotFocus()
  With txtEvalTo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEvalTo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtRating_GotFocus(Index As Integer)
  With txtRating(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtRating_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtRating_LostFocus(Index As Integer)
  
  Dim i As Integer
  
  With txtRating(Index)
    lblScore(Index).Caption = CInt(lblMultiplier(Index).Caption) * CInt(.Text)
  End With
  
  lblTtlScore.Caption = 0
  
  For i = lblScore.LBound To lblScore.UBound
    lblTtlScore.Caption = CInt(lblTtlScore.Caption) + CInt(lblScore(i).Caption)
  Next
  
End Sub

Private Sub txtJobTitle_GotFocus()
  With txtJobTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtJobTitle_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtRemarks_GotFocus()
  With txtRemarks
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtRemarks_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Clear_Data()
  
  Dim i               As Integer
  
  mAdd = True
  txtEmployee.Text = ""
  txtEmployee.Tag = ""
  txtJobTitle.Text = ""
  txtJobTitle.Tag = ""
  txtDivision.Text = ""
  txtDivision.Tag = ""
  txtCostCenter.Text = ""
  txtCostCenter.Tag = ""
  txtDateHired.Text = ""
  txtEvalDate.Text = ""
  txtEvalFrom.Text = ""
  txtEvalTo.Text = ""
  tdbEvalType.BoundText = ""
  
  For i = txtRating.LBound To txtRating.UBound
    txtRating(i).Text = 1
    txtRating_LostFocus i
  Next
  
  txtRemarks.Text = ""
  
  cmdEmployee.Enabled = True
  cmdJobTitle.Enabled = True
  cmdDivision.Enabled = True
  cmdCostCenter.Enabled = True

End Sub

Private Sub txtEvalDate_GotFocus()
  With txtEvalDate
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtEvalDate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub


