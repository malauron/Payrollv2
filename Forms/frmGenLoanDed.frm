VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmGenLoanDed 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      Height          =   705
      Left            =   75
      TabIndex        =   5
      Top             =   1050
      Width           =   12720
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   315
         Left            =   1530
         TabIndex        =   6
         Tag             =   "Municipal"
         Top             =   255
         Width           =   3900
         _ExtentX        =   6879
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
         Columns(0).DataField=   "provcode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   "provdesc"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "from"
         Columns(2).DataField=   "wrkdatefrom"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "To"
         Columns(3).DataField=   "wrkdateto"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1879"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2355"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2249"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2752"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
         EditFont        =   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _PropDict       =   $"frmGenLoanDed.frx":0000
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H404040&"
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Period"
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
         Height          =   225
         Left            =   255
         TabIndex        =   7
         Top             =   300
         Width           =   1830
      End
   End
   Begin TDBNumber6Ctl.TDBNumber txtamtded 
      Height          =   270
      Left            =   4920
      TabIndex        =   4
      Top             =   1740
      Visible         =   0   'False
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   476
      Calculator      =   "frmGenLoanDed.frx":00AA
      Caption         =   "frmGenLoanDed.frx":00CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmGenLoanDed.frx":0136
      Keys            =   "frmGenLoanDed.frx":0154
      Spin            =   "frmGenLoanDed.frx":019E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   2
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
      MinValue        =   -9999
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
   Begin VB.Frame fra2 
      BackColor       =   &H00F6F8F8&
      Height          =   525
      Left            =   75
      TabIndex        =   1
      Top             =   5850
      Width           =   6765
      Begin TDBText6Ctl.TDBText txtSearch 
         Height          =   315
         Left            =   2250
         TabIndex        =   2
         Top             =   150
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   556
         Caption         =   "frmGenLoanDed.frx":01C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmGenLoanDed.frx":0232
         Key             =   "frmGenLoanDed.frx":0250
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
      Begin VB.Label Label2 
         BackColor       =   &H00F6F8F8&
         Caption         =   "Search employee name"
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
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   2250
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgGenLoanDed 
      Height          =   3060
      Left            =   255
      TabIndex        =   0
      Top             =   2070
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   5398
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "empno"
      Columns(0).DataField=   "empno"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Employee name"
      Columns(1).DataField=   "fullname"
      Columns(1).NumberFormat=   "MM/DD/YYYY HH:NN:SS AM/PM"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "loanno"
      Columns(2).DataField=   "loanno"
      Columns(2).NumberFormat=   "MM/DD/YYYY HH:NN:SS AM/PM"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Loan Types"
      Columns(3).DataField=   "loantypes"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Deductable Amount"
      Columns(4).DataField=   "amtded"
      Columns(4).NumberFormat=   "#,##0.00"
      Columns(4).ExternalEditor=   "txtamtded"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "balance"
      Columns(5).DataField=   "balance"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "ttlamtpaid"
      Columns(6).DataField=   "ttlamtpaid"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=6932"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6853"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1693"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1614"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=6747"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=6668"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=8196"
      Splits(0)._ColumnProps(38)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(40)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(41)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(43)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=8196"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
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
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   16185592
      RowDividerColor =   16185592
      RowSubDividerColor=   16185592
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF6F8F8&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H400000&"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HF6F8F8&"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H0&,.fgcolor=&H0&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HEBFEEB&,.fgcolor=&H0&"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H0&"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(17)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFF0EA&"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HF6F8F8&"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HF6F8F8&"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
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
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.locked=-1"
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
   Begin CitronSoftwarePayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   120
      TabIndex        =   8
      Top             =   435
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      BorderColor     =   14215660
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   420
         Left            =   1230
         TabIndex        =   9
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Save"
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
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   420
         Left            =   75
         TabIndex        =   10
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Generate"
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
   End
   Begin CitronSoftwarePayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Generate loan deductions"
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
      ForeColor       =   3186872
      GradTheme       =   2
   End
End
Attribute VB_Name = "frmGenLoanDed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsLoanDed       As ADODB.Recordset
Dim rsTmpLoanded    As ADODB.Recordset

Private Sub cmdGenerate_Click()

    Dim rsBalance   As ADODB.Recordset

    If Trim(tdbPayrollPeriod.Text) = "" Or IsNull(tdbPayrollPeriod.SelectedItem) Or tdbPayrollPeriod.ApproxCount = 0 Then
        MsgBox "Please choose a payroll period before generating a loan deduction.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus
        Exit Sub
    End If
    
    Create_TmpLoanded
    
    With rsTmpLoanded
        NetOpen rsLoanDed, "", "select x1.*, concat(x2.lastname,', ',x2.firstname,' ',x2.middlename ) fullname,x3.loantypes from loans x1 " & _
                               "left outer join employee x2 on x1.empno = x2.empno " & _
                               "left outer join loantypes x3 on x1.loantypescode = x3.loantypescode " & _
                               "where x1.status = 'Active' and x2.payfreqcode = '" & tdbPayrollPeriod.Columns("payfreqcode").Text & "'"
        
        If rsLoanDed.RecordCount > 0 Then
            rsLoanDed.MoveFirst
            Do While Not rsLoanDed.EOF
                
                .AddNew
                .Fields("empno") = rsLoanDed!empno
                .Fields("fullname") = rsLoanDed!Fullname
                .Fields("loanno") = rsLoanDed!loanno
                .Fields("loantypescode") = rsLoanDed!loantypescode
                .Fields("loantypes") = rsLoanDed!LoanTypes
                .Fields("amtded") = rsLoanDed!dedperpayday
                
                NetOpen rsBalance, "", "select balance from loanded " & _
                                       "where loanno = '" & rsLoanDed!loanno & "'  and fnlz = 'Y' " & _
                                       "order by dateposted desc limit 1"
                                       
                If rsBalance.RecordCount > 0 Then
                    .Fields("balance") = rsBalance!balance
                    .Fields("ttlamtpaid") = rsBalance!ttlamtpaid
                Else
                    .Fields("balance") = rsLoanDed!loanamnt
                    .Fields("ttlamtpaid") = 0
                End If
                
                .Update
                rsLoanDed.MoveNext
                DoEvents
            Loop
        End If
    
        NetOpen rsLoanDed, "", "select * from loanded where percode = '" & tdbPayrollPeriod.BoundText & "'"
        
        If rsLoanDed.RecordCount > 0 Then
            rsLoanDed.MoveFirst
            Do While Not rsLoanDed.EOF
                .MoveFirst
                .Find "loanno = '" & rsLoanDed!loanno & "'"
                .Fields("amtded") = rsLoanDed!amtded
                rsLoanDed.MoveNext
                DoEvents
            Loop
        End If
    End With
    
End Sub

Private Sub cmdSave_Click()
    With rsTmpLoanded
        If .RecordCount > 0 Then
            .MoveFirst
            CitronPayroll.Execute "set autocommit = 0"
            CitronPayroll.BeginTrans
            CitronPayroll.Execute "delete from loanded where percode = '" & tdbPayrollPeriod.BoundText & "'"
            Do While Not .EOF
                CitronPayroll.Execute "insert into loanded(loanno,loantypescode,empno,percode,payyear,paymonth,amtded,dateposted,ttlamtpaid,balance,fnlz) values " & _
                    "('" & !loanno & "','" & !loantypescode & "','" & !empno & "','" & tdbPayrollPeriod.BoundText & "','" & tdbPayrollPeriod.Columns("payyear") & "','" & tdbPayrollPeriod.Columns("paymonth") & "', " & _
                    "" & !amtded & ",'" & Format(Now, "YYYY-MM-DD") & "'," & CDbl(!ttlamtpaid) + CDbl(!amtded) & "," & CDbl(!balance) - CDbl(!amtded) & ",'N')"
                .MoveNext
            Loop
            CitronPayroll.CommitTrans
        End If
    End With
End Sub

Private Sub Form_Load()
    Add_MDIButton Me.Name, titlebar.Caption
    bind_tdb CitronPayroll, tdbPayrollPeriod, "select percode, description,wrkdatefrom,wrkdateto,payfreqcode,paymonth,payyear from payrollperiod", "description", "percode"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    titlebar.Move 0, 0, Me.ScaleWidth
    
    With fraButton
        .Top = titlebar.Top + titlebar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With fra1
        .Top = fraButton.Height + fraButton.Top
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    
    With fra2
        .Top = Me.ScaleHeight - .Height
        .Left = 0
    End With
    
    With tdgGenLoanDed
        .Top = fra1.Top + fra1.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - (.Top + fra2.Height)
    End With
    
End Sub

Private Sub Create_TmpLoanded()

    Set rsTmpLoanded = Nothing
    Set rsTmpLoanded = New ADODB.Recordset
    
    With rsTmpLoanded
        .Fields.Append "empno", adVarChar, 15
        .Fields.Append "fullname", adVarChar, 70
        .Fields.Append "loanno", adVarChar, 7
        .Fields.Append "loantypes", adVarChar, 70
        .Fields.Append "loantypescode", adVarChar, 7
        .Fields.Append "amtded", adDouble
        .Fields.Append "balance", adDouble
        .Fields.Append "ttlamtpaid", adDouble
        .Open
        .Sort = "fullname"
    End With
        
    Set tdgGenLoanDed.DataSource = rsTmpLoanded
    
End Sub

Private Sub tdgGenLoanDed_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    With tdgGenLoanDed
        If ColIndex = .Columns("amtded").ColIndex Then
            If CDbl(.Columns("amtded").Text) > CDbl(.Columns("balance").Text) Then
                MsgBox "The amount you entered is greater than the remaining balance of the employee's loan.", vbExclamation + vbOKOnly
                Cancel = True
                .SetFocus
                .Col = ColIndex
            End If
        End If
    End With
End Sub

