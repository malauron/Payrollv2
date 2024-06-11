VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmParameter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11415
   StartUpPosition =   1  'CenterOwner
   Begin LinkProPayroll.b8ChildTitleBar titlebar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
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
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   -15
      TabIndex        =   1
      Top             =   255
      Width           =   11460
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pull-out rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   945
         Left            =   5790
         TabIndex        =   46
         Top             =   5790
         Width           =   5550
         Begin TDBNumber6Ctl.TDBNumber txtPullOutRate 
            Height          =   300
            Left            =   4065
            TabIndex        =   47
            Top             =   225
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0000
            Caption         =   "frmParameter.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":008C
            Keys            =   "frmParameter.frx":00AA
            Spin            =   "frmParameter.frx":00F4
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
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
            MaxValue        =   999999999999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Daily pull-out rate"
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
            Height          =   195
            Left            =   195
            TabIndex        =   48
            Top             =   285
            Width           =   2820
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Late, Overtime and Pull-out parameters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5640
         Left            =   5790
         TabIndex        =   39
         Top             =   135
         Width           =   5550
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ovetime Break Allowance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   3615
            Left            =   90
            TabIndex        =   53
            Top             =   1935
            Width           =   5370
            Begin TrueOleDBGrid80.TDBGrid tdgOtBreakAllowance 
               Height          =   3360
               Left            =   45
               TabIndex        =   54
               Top             =   225
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   5927
               _LayoutType     =   4
               _RowHeight      =   16
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "No. Of Hours"
               Columns(0).DataField=   "noofhrs"
               Columns(0).NumberFormat=   "#,##0"
               Columns(0).ExternalEditor=   "txtNoOfHrs"
               Columns(0).ExternalEditor.vt=   8
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Break Allowance In Minutes"
               Columns(1).DataField=   "allowablemin"
               Columns(1).NumberFormat=   "#,##0"
               Columns(1).ExternalEditor=   "txtAllowableMin"
               Columns(1).ExternalEditor.vt=   8
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
               Splits(0)._ColumnProps(1)=   "Column(0).Width=4736"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4657"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=514"
               Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
            Begin TDBNumber6Ctl.TDBNumber txtAllowableMin 
               Height          =   300
               Left            =   3690
               TabIndex        =   55
               Top             =   615
               Visible         =   0   'False
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   529
               Calculator      =   "frmParameter.frx":011C
               Caption         =   "frmParameter.frx":013C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmParameter.frx":01A8
               Keys            =   "frmParameter.frx":01C6
               Spin            =   "frmParameter.frx":0210
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   1
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#,###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#,###,###,###,##0"
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
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   27590657
               Value           =   0
               MaxValueVT      =   1145962501
               MinValueVT      =   1414463493
            End
            Begin TDBNumber6Ctl.TDBNumber txtNoOfHrs 
               Height          =   300
               Left            =   390
               TabIndex        =   56
               Top             =   585
               Visible         =   0   'False
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   529
               Calculator      =   "frmParameter.frx":0238
               Caption         =   "frmParameter.frx":0258
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmParameter.frx":02C4
               Keys            =   "frmParameter.frx":02E2
               Spin            =   "frmParameter.frx":032C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   1
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#,###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#,###,###,###,##0"
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
               ReadOnly        =   0
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   27590657
               Value           =   0
               MaxValueVT      =   1145962501
               MinValueVT      =   1414463493
            End
         End
         Begin TDBNumber6Ctl.TDBNumber txtHrsPerDay 
            Height          =   300
            Left            =   4065
            TabIndex        =   40
            Top             =   225
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0354
            Caption         =   "frmParameter.frx":0374
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":03E0
            Keys            =   "frmParameter.frx":03FE
            Spin            =   "frmParameter.frx":0448
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
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
            MaxValue        =   999999999999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtLateAllowance 
            Height          =   300
            Left            =   4065
            TabIndex        =   43
            Top             =   570
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0470
            Caption         =   "frmParameter.frx":0490
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":04FC
            Keys            =   "frmParameter.frx":051A
            Spin            =   "frmParameter.frx":0564
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtOTAllow 
            Height          =   300
            Left            =   4065
            TabIndex        =   44
            Top             =   915
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":058C
            Caption         =   "frmParameter.frx":05AC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0618
            Keys            =   "frmParameter.frx":0636
            Spin            =   "frmParameter.frx":0680
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtUTmin 
            Height          =   300
            Left            =   4065
            TabIndex        =   51
            Top             =   1260
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":06A8
            Caption         =   "frmParameter.frx":06C8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0734
            Keys            =   "frmParameter.frx":0752
            Spin            =   "frmParameter.frx":079C
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label21 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "No. of minutes to be considered undtertime"
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
            Height          =   345
            Left            =   195
            TabIndex        =   52
            Top             =   1320
            Width           =   3840
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum minutes for overtime:"
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
            Height          =   345
            Left            =   195
            TabIndex        =   45
            Top             =   975
            Width           =   3840
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Late alowance in minutes:"
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
            Height          =   195
            Left            =   195
            TabIndex        =   42
            Top             =   630
            Width           =   2820
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "No. of working hours per day:"
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
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   285
            Width           =   2820
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Computing Overtime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   5550
         Begin TDBNumber6Ctl.TDBNumber txtotregprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   5
            Top             =   225
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":07C4
            Caption         =   "frmParameter.frx":07E4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0850
            Keys            =   "frmParameter.frx":086E
            Spin            =   "frmParameter.frx":08B8
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotrstprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   7
            Top             =   570
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":08E0
            Caption         =   "frmParameter.frx":0900
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":096C
            Keys            =   "frmParameter.frx":098A
            Spin            =   "frmParameter.frx":09D4
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotspcprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   37
            Top             =   915
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":09FC
            Caption         =   "frmParameter.frx":0A1C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0A88
            Keys            =   "frmParameter.frx":0AA6
            Spin            =   "frmParameter.frx":0AF0
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotlegprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   38
            Top             =   1260
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0B18
            Caption         =   "frmParameter.frx":0B38
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0BA4
            Keys            =   "frmParameter.frx":0BC2
            Spin            =   "frmParameter.frx":0C0C
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Legal holiday:"
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
            Height          =   300
            Left            =   195
            TabIndex        =   36
            Top             =   1320
            Width           =   3915
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Special holiday:"
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
            Height          =   300
            Left            =   195
            TabIndex        =   35
            Top             =   960
            Width           =   3915
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On a restday:"
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
            Height          =   495
            Left            =   180
            TabIndex        =   6
            Top             =   615
            Width           =   3915
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On ordinary days: "
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
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   300
            Width           =   2235
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Computing Night Premium Overtime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1680
         Left            =   90
         TabIndex        =   22
         Top             =   5790
         Width           =   5550
         Begin TDBNumber6Ctl.TDBNumber txtotniteregprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   23
            Top             =   225
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0C34
            Caption         =   "frmParameter.frx":0C54
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0CC0
            Keys            =   "frmParameter.frx":0CDE
            Spin            =   "frmParameter.frx":0D28
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotniterstprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   24
            Top             =   570
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0D50
            Caption         =   "frmParameter.frx":0D70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0DDC
            Keys            =   "frmParameter.frx":0DFA
            Spin            =   "frmParameter.frx":0E44
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotnitespcprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   29
            Top             =   915
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0E6C
            Caption         =   "frmParameter.frx":0E8C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":0EF8
            Keys            =   "frmParameter.frx":0F16
            Spin            =   "frmParameter.frx":0F60
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtotnitelegprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   30
            Top             =   1260
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":0F88
            Caption         =   "frmParameter.frx":0FA8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":1014
            Keys            =   "frmParameter.frx":1032
            Spin            =   "frmParameter.frx":107C
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Legal holiday:"
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
            Height          =   300
            Left            =   195
            TabIndex        =   28
            Top             =   1335
            Width           =   3915
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Special holiday:"
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
            Height          =   300
            Left            =   195
            TabIndex        =   27
            Top             =   975
            Width           =   3915
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On ordinary day:"
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
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   285
            Width           =   3900
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On rest day:"
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
            Height          =   300
            Left            =   195
            TabIndex        =   25
            Top             =   615
            Width           =   3915
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Computing Night Premiums"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1680
         Left            =   90
         TabIndex        =   17
         Top             =   1980
         Width           =   5550
         Begin TDBNumber6Ctl.TDBNumber txtniteregprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   18
            Top             =   225
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":10A4
            Caption         =   "frmParameter.frx":10C4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":1130
            Keys            =   "frmParameter.frx":114E
            Spin            =   "frmParameter.frx":1198
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtniterstprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   19
            Top             =   570
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":11C0
            Caption         =   "frmParameter.frx":11E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":124C
            Keys            =   "frmParameter.frx":126A
            Spin            =   "frmParameter.frx":12B4
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtnitespcprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   33
            Top             =   915
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":12DC
            Caption         =   "frmParameter.frx":12FC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":1368
            Keys            =   "frmParameter.frx":1386
            Spin            =   "frmParameter.frx":13D0
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtnitelegprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   34
            Top             =   1260
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":13F8
            Caption         =   "frmParameter.frx":1418
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":1484
            Keys            =   "frmParameter.frx":14A2
            Spin            =   "frmParameter.frx":14EC
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Legal holiday:"
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
            Height          =   300
            Left            =   180
            TabIndex        =   32
            Top             =   1335
            Width           =   3915
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On Special holiday:"
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
            Height          =   300
            Left            =   180
            TabIndex        =   31
            Top             =   975
            Width           =   3915
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On rest day:"
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
            Height          =   495
            Left            =   180
            TabIndex        =   21
            Top             =   615
            Width           =   3915
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "On ordinary day:"
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
            Height          =   195
            Left            =   195
            TabIndex        =   20
            Top             =   285
            Width           =   3900
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Computig pay for work done on "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2085
         Left            =   90
         TabIndex        =   8
         Top             =   3690
         Width           =   5550
         Begin TDBNumber6Ctl.TDBNumber txtspcholprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   9
            Top             =   540
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":1514
            Caption         =   "frmParameter.frx":1534
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":15A0
            Keys            =   "frmParameter.frx":15BE
            Spin            =   "frmParameter.frx":1608
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtrestspcholprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   10
            Top             =   885
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":1630
            Caption         =   "frmParameter.frx":1650
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":16BC
            Keys            =   "frmParameter.frx":16DA
            Spin            =   "frmParameter.frx":1724
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtlegholprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   13
            Top             =   1230
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":174C
            Caption         =   "frmParameter.frx":176C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":17D8
            Keys            =   "frmParameter.frx":17F6
            Spin            =   "frmParameter.frx":1840
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtrestlegholprct 
            Height          =   300
            Left            =   4065
            TabIndex        =   14
            Top             =   1575
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":1868
            Caption         =   "frmParameter.frx":1888
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":18F4
            Keys            =   "frmParameter.frx":1912
            Spin            =   "frmParameter.frx":195C
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin TDBNumber6Ctl.TDBNumber txtRstPrct 
            Height          =   300
            Left            =   4065
            TabIndex        =   49
            Top             =   195
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   529
            Calculator      =   "frmParameter.frx":1984
            Caption         =   "frmParameter.frx":19A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmParameter.frx":1A10
            Keys            =   "frmParameter.frx":1A2E
            Spin            =   "frmParameter.frx":1A78
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   1
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00 %"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##0.00 %"
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
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1145962501
            MinValueVT      =   1414463493
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Rest day:"
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
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   270
            Width           =   3900
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Legal holiday:"
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
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1290
            Width           =   3900
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Legal holiday, which is also a rest day."
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
            Height          =   495
            Left            =   180
            TabIndex        =   15
            Top             =   1620
            Width           =   3915
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Special holiday:"
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
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   615
            Width           =   3900
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Special day, which is also a  rest day:"
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
            Height          =   495
            Left            =   180
            TabIndex        =   11
            Top             =   930
            Width           =   3915
         End
      End
      Begin lvButton.lvButtons_H cmdUpdate 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   9330
         TabIndex        =   2
         Top             =   7005
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   688
         Caption         =   "&Update"
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
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsParmtr                As ADODB.Recordset
Dim rsOtBreakAllowance      As ADODB.Recordset
Dim rsOtBreakAllowanceTmp   As ADODB.Recordset

Private Sub cmdUpdate_Click()

    Dim rsTmp               As ADODB.Recordset

    Set rsTmp = New ADODB.Recordset
    Set rsTmp = rsOtBreakAllowanceTmp
    
    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Not IsNumeric(!noofhrs) Then
                    MsgBox "Check for invalid number of hours.", vbExclamation + vbOKOnly
                    tdgOtBreakAllowance.SetFocus
                    Exit Sub
                End If
                If Not IsNumeric(!allowablemin) Then
                    MsgBox "Check for invalid number of minute/s.", vbExclamation + vbOKCancel
                    tdgOtBreakAllowance.SetFocus
                    Exit Sub
                End If
                .MoveNext
            Loop
        End If
    End With

    
    ConMain.BeginTrans
    
    If rsParmtr.RecordCount = 0 Then
        ConMain.Execute "insert into parmtr(otregprct,otrstprct,otspcprct,otlegprct,rstprct,spcholprct,restspcholprct,legholprct, " & _
            "restlegholprct,niteregprct,niterstprct,nitespcprct,nitelegprct,otniteregprct,otniterstprct,otnitespcprct,otnitelegprct,hrsperday,lateallowance,OTAllowance,utmin,pulloutrate) values " & _
            "(" & Format(txtotregprct, "##0.00") & "," & Format(txtotrstprct, "##0.00") & "," & Format(txtotspcprct, "##0.00") & "," & Format(txtotlegprct, "##0.00") & "," & Format(txtRstPrct, "##0.00") & "," & Format(txtspcholprct, "##0.00") & "," & Format(txtrestspcholprct, "##0.00") & "," & Format(txtlegholprct, "##0.00") & _
            "," & Format(txtrestlegholprct, "##0.00") & "," & Format(txtniteregprct, "##0.00") & "," & Format(txtniterstprct, "##0.00") & "," & Format(txtnitespcprct, "##0.00") & "," & Format(txtnitelegprct, "##0.00") & "," & Format(txtotniteregprct, "##0.00") & "," & Format(txtotniterstprct, "##0.00") & "," & Format(txtotnitespcprct, "##0.00") & "," & Format(txtotnitelegprct, "##0.00") & ", " & _
            Format(txtHrsPerDay, "#,##0.00") & "," & Format(txtLateAllowance, "##0") & "," & Format(txtOTAllow, "##0") & "," & Format(txtUTmin, "##0") & "," & Format(txtPullOutRate, "#,##0.00") & ")"
    Else
        ConMain.Execute "update parmtr set otregprct =" & Format(txtotregprct, "##0.00") & " ,otrstprct = " & Format(txtotrstprct, "##0.00") & ",otspcprct = " & Format(txtotspcprct, "##0.00") & ",otlegprct = " & Format(txtotlegprct, "##0.00") & ", " & _
            "rstprct = " & Format(txtRstPrct, "##0.00") & ",spcholprct = " & Format(txtspcholprct, "##0.00") & ",restspcholprct = " & Format(txtrestspcholprct, "##0.00") & ",legholprct= " & Format(txtlegholprct, "##0.00") & _
            ",restlegholprct = " & Format(txtrestlegholprct, "##0.00") & ",niteregprct = " & Format(txtniteregprct, "##0.00") & ",niterstprct = " & Format(txtniterstprct, "##0.00") & ",nitespcprct = " & Format(txtnitespcprct, "##0.00") & ",nitelegprct = " & Format(txtnitelegprct, "##0.00") & ", " & _
            "otniteregprct = " & Format(txtotniteregprct, "##0.00") & ",otniterstprct = " & Format(txtotniterstprct, "##0.00") & ",otnitespcprct = " & Format(txtotnitespcprct, "##0.00") & ",otnitelegprct = " & Format(txtotnitelegprct, "##0.00") & ", " & _
            "hrsperday = " & Format(txtHrsPerDay, "##0.00") & ",lateallowance = " & Format(txtLateAllowance, "##0") & ", OTAllowance = " & Format(txtOTAllow, "##0") & ", utmin = " & Format(txtUTmin, "##0") & ",pulloutrate = " & Format(txtPullOutRate, "##0.00") & ""
    End If
    
    
    ConMain.Execute "Delete from otbreakallowance"
    
    With rsOtBreakAllowanceTmp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                ConMain.Execute "insert into otbreakallowance(noofhrs,allowablemin) values (" & !noofhrs & "," & !allowablemin & ")"
                .MoveNext
            Loop
        End If
    End With
    
    ConMain.CommitTrans
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    NetOpen rsParmtr, "select * from parmtr"
    
    With rsParmtr
        If .RecordCount > 0 Then
            txtotregprct.Text = Format(!otregprct / 100, "##0.00 %")
            txtotrstprct.Text = Format(!otrstprct / 100, "##0.00 %")
            txtotspcprct.Text = Format(!otspcprct / 100, "##0.00 %")
            txtotlegprct.Text = Format(!otlegprct / 100, "##0.00 %")
            txtspcholprct.Text = Format(!spcholprct / 100, "##0.00 %")
            txtRstPrct.Text = Format(!rstprct / 100, "##0.00 %")
            txtrestspcholprct.Text = Format(!restspcholprct / 100, "##0.00 %")
            txtlegholprct.Text = Format(!legholprct / 100, "##0.00 %")
            txtrestlegholprct.Text = Format(!restlegholprct / 100, "##0.00 %")
            txtniteregprct.Text = Format(!niteregprct / 100, "##0.00 %")
            txtniterstprct.Text = Format(!niterstprct / 100, "##0.00 %")
            txtnitespcprct.Text = Format(!nitespcprct / 100, "##0.00 %")
            txtnitelegprct.Text = Format(!nitelegprct / 100, "##0.00 %")
            txtotniteregprct.Text = Format(!otniteregprct / 100, "##0.00 %")
            txtotniterstprct.Text = Format(!otniterstprct / 100, "##0.00 %")
            txtotnitespcprct.Text = Format(!otnitespcprct / 100, "##0.00 %")
            txtotnitelegprct.Text = Format(!otnitelegprct / 100, "##0.00 %")
            txtHrsPerDay.Text = Format(!hrsperday, "#,##0.00")
            txtLateAllowance.Text = Format(!lateallowance, "#,##0")
            txtOTAllow.Text = Format(!otallowance, "#,##0")
            txtUTmin.Text = Format(!utmin, "#,##0")
            txtPullOutRate.Text = Format(!pulloutrate, "#,##0.00")
        End If
    End With
    
    Create_TmpOtBreakAllowance
    
    NetOpen rsOtBreakAllowance, "select * from otbreakallowance"
    
    With rsOtBreakAllowance
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                rsOtBreakAllowanceTmp.AddNew
                rsOtBreakAllowanceTmp.Fields("noofhrs") = !noofhrs
                rsOtBreakAllowanceTmp.Fields("allowablemin") = !allowablemin
                rsOtBreakAllowanceTmp.Update
                .MoveNext
            Loop
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    TitleBar.Move 0, 0, Me.ScaleWidth
End Sub

Private Sub Create_TmpOtBreakAllowance()

    Set rsOtBreakAllowanceTmp = Nothing
    Set rsOtBreakAllowanceTmp = New ADODB.Recordset
    
    With rsOtBreakAllowanceTmp
        .Fields.Append "noofhrs", adInteger
        .Fields.Append "allowablemin", adInteger
        .Open
    End With

    With tdgOtBreakAllowance
        Set .DataSource = rsOtBreakAllowanceTmp
        .ReBind
        .Refresh
        .ReOpen
    End With
    
End Sub

Private Sub tdgotbreakallowance_BeforeDelete(Cancel As Integer)
    If MsgBox("Do you want to delete this entry?", vbQuestion + vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub tdgOtBreakAllowance_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Error_Hndlr
    
    With tdgOtBreakAllowance
        If .ApproxCount > 0 Then
            If Not .EOF Then
                If KeyCode = 46 Then
                    If txtNoOfHrs.Visible = False And txtAllowableMin.Visible = False Then
                      .Delete
                      tdgOtBreakAllowance.SetFocus
                    End If
                End If
            End If
        End If
    End With
    
Exit Sub
    
Error_Hndlr:

    If err.Number = -2147467259 Then
        MsgBox "Deletion process was aborted.", vbExclamation + vbOKOnly
        tdgOtBreakAllowance.SetFocus
    Else
        MsgBox "Unexpected error occured. " & vbCrLf & "Error number: " & err.Number & vbCrLf & "Error Description: " & err.Description
    End If
    
End Sub

Private Sub txtNoOfhrs_LostFocus()
    tdgOtBreakAllowance.SetFocus
End Sub

Private Sub txtAllowableMin_LostFocus()
    tdgOtBreakAllowance.SetFocus

End Sub
