VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmRptMealAgreementForm 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9780
   WindowState     =   2  'Maximized
   Begin VB.Frame fraParmtr 
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      Height          =   5880
      Left            =   1530
      TabIndex        =   0
      Top             =   1410
      Width           =   3465
      Begin OsenXPCntrl.OsenXPButton cmdView 
         Height          =   435
         Left            =   495
         TabIndex        =   1
         Top             =   4905
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   767
         BTYPE           =   5
         TX              =   "&View"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16185592
         BCOLO           =   16185592
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmRptMealAgreementForm.frx":0000
         PICN            =   "frmRptMealAgreementForm.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBText6Ctl.TDBText txtEmployee 
         Height          =   300
         Left            =   75
         TabIndex        =   4
         Top             =   345
         Width           =   3030
         _Version        =   65536
         _ExtentX        =   5345
         _ExtentY        =   529
         Caption         =   "frmRptMealAgreementForm.frx":05B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRptMealAgreementForm.frx":0622
         Key             =   "frmRptMealAgreementForm.frx":0640
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
      Begin lvButton.lvButtons_H cmdEmployee 
         Height          =   315
         Left            =   3090
         TabIndex        =   5
         Top             =   345
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
      Begin TDBText6Ctl.TDBText txtJobTitle 
         Height          =   300
         Left            =   75
         TabIndex        =   7
         Top             =   1140
         Width           =   3030
         _Version        =   65536
         _ExtentX        =   5345
         _ExtentY        =   529
         Caption         =   "frmRptMealAgreementForm.frx":0684
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRptMealAgreementForm.frx":06F0
         Key             =   "frmRptMealAgreementForm.frx":070E
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
      Begin lvButton.lvButtons_H cmdJobTitle 
         Height          =   315
         Left            =   3090
         TabIndex        =   8
         Top             =   1140
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
         Bindings        =   "frmRptMealAgreementForm.frx":0752
         DataMember      =   "tdbJob"
         Height          =   300
         Left            =   30
         TabIndex        =   10
         Top             =   2610
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
         _PropDict       =   $"frmRptMealAgreementForm.frx":0763
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Title"
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
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   9
         Top             =   930
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
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
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   6
         Top             =   135
         Width           =   1725
      End
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer xrpt 
      Height          =   5790
      Left            =   2730
      TabIndex        =   2
      Top             =   -15
      Width           =   6405
      lastProp        =   600
      _cx             =   11298
      _cy             =   10213
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   225
      TabIndex        =   3
      Top             =   45
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   609
      Caption         =   "Meal Agreement Form"
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
End
Attribute VB_Name = "frmRptMealAgreementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mVoucherValidPeriodCode     As Integer

Dim CrxRep                      As CRAXDRT.Report
Dim mTxt                        As TDBText

Private Sub cmdEmployee_Click()
  bind_tdb ConMain, tdbTmp, "select employeecode,concat(lastname,',',firstname,' ',middlename) fullname from employee order by lastname,firstname,middlename", "fullname", "employeecode"
  Set mTxt = txtEmployee
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
End Sub

Private Sub cmdJobTitle_Click()
  bind_tdb ConMain, tdbTmp, "select jobtitlecode,description from jobtitle order by description", "description", "jobtitlecode"
  Set mTxt = txtJobTitle
  tdbTmp.Move mTxt.Left, mTxt.Top, mTxt.Width, mTxt.Height
  tdbTmp.BoundText = mTxt.Tag
  mTxt.Visible = False
  tdbTmp.Visible = True
  tdbTmp.SetFocus
  SendKeys "{F4}"
End Sub

Private Sub cmdView_Click()

    Dim CrxApp                  As CRAXDRT.Application
    Dim crxDatabase             As CRAXDRT.Database
    Dim crxDatabaseTable        As CRAXDRT.DatabaseTable
    Dim crxDatabaseTables       As CRAXDRT.DatabaseTables
    
    Dim isIn                    As Boolean

      
    Set CrxApp = Nothing
    Set CrxRep = Nothing
    
    Set CrxApp = New CRAXDRT.Application
    Set CrxRep = New CRAXDRT.Report
    
    Set CrxRep = CrxApp.OpenReport(App.Path & "\reports\rptMealAgreementForm.rpt")
    
    Set crxDatabase = CrxRep.Database
    Set crxDatabaseTables = crxDatabase.Tables
    
    For Each crxDatabaseTable In crxDatabaseTables
        crxDatabaseTable.ConnectionProperties.DeleteAll
        crxDatabaseTable.ConnectionProperties.Add "connection string", ""
        crxDatabaseTable.ConnectionProperties("connection string") = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & SQLServerName & "; DATABASE=" & SQLDatabase & "; UID=" & SQLUsername & "; PWD=" & SQLPassword & "; PORT='" & SQLPort & "'"
    Next crxDatabaseTable
    
    CrxRep.ParameterFields.GetItemByName("employeecode").AddCurrentValue CInt(txtEmployee.Tag)
    'CrxRep.ParameterFields.GetItemByName("mDate").AddCurrentValue CDate(Format(txtDate.Value, "MM/DD/YYYY"))

    xrpt.ReportSource = CrxRep
    xrpt.ViewReport
    xrpt.Zoom 100

    Set crxDatabase = Nothing
    Set crxDatabaseTable = Nothing
    Set crxDatabaseTables = Nothing
    Set CrxApp = Nothing
    
End Sub

Private Sub Form_Load()

    Dim rsTmp           As ADODB.Recordset
    
    Dim i               As Integer

    Add_MDIButton Me.Name, TitleBar.Caption
    
    Set rsTmp = Nothing
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    TitleBar.Move 0, 0, Me.ScaleWidth

    With fraParmtr
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Height = Me.ScaleHeight - .Top
    End With

    With xrpt
        .Top = TitleBar.Top + TitleBar.Height
        .Left = fraParmtr.Left + fraParmtr.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top
    End With
    
End Sub

Private Sub tdbTmp_KeyPress(KeyAscii As Integer)
  Dim rs  As ADODB.Recordset
  If KeyAscii = 13 Then
    If Trim(tdbTmp.Text) <> "" And Not IsNull(tdbTmp.SelectedItem) And tdbTmp.ApproxCount > 0 Then
      If mTxt.Name = "txtEmployee" Then
        mTxt.Tag = tdbTmp.BoundText
        mTxt.Text = tdbTmp.Text
        NetOpen rs, "select x1.jobtitlecode,x2.description from employee x1 " & _
                    "left outer join jobtitle x2 on x1.jobtitlecode=x2.jobtitlecode " & _
                    "where x1.jobtitlecode=(select jobtitlecode from employee where employeecode = " & mTxt.Tag & ")"
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          txtJobTitle.Tag = rs!jobtitlecode
          txtJobTitle.Text = rs!Description
        End If
      Else
        If MsgBox("Confirm job description update.", vbOKCancel) = vbOK Then
          mTxt.Tag = tdbTmp.BoundText
          mTxt.Text = tdbTmp.Text
          ConMain.Execute "update employee set jobtitlecode = " & mTxt.Tag & " where employeecode=" & txtEmployee.Tag & ""
        End If
      End If
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

Private Sub xrpt_PrintButtonClicked(UseDefault As Boolean)
    CrxRep.PrinterSetup Me.hwnd
End Sub


