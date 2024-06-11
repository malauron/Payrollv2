VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmADOvertime2 
   BackColor       =   &H00F6F8F8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Details"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   -15
      TabIndex        =   7
      Top             =   1995
      Width           =   7935
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   60
         TabIndex        =   3
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
         Image           =   "frmADOvertime2.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   1800
         TabIndex        =   2
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
         Image           =   "frmADOvertime2.frx":0CDA
         cBack           =   14737632
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   2070
      Left            =   30
      TabIndex        =   4
      Top             =   -75
      Width           =   7815
      Begin TDBNumber6Ctl.TDBNumber txtOtHrs 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   585
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   556
         Calculator      =   "frmADOvertime2.frx":19B4
         Caption         =   "frmADOvertime2.frx":19D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADOvertime2.frx":1A40
         Keys            =   "frmADOvertime2.frx":1A5E
         Spin            =   "frmADOvertime2.frx":1AA8
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
         ForeColor       =   4210752
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
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin TrueOleDBList80.TDBCombo tdbEmployee 
         Height          =   345
         Left            =   1680
         TabIndex        =   0
         Tag             =   "Municipal"
         Top             =   180
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   609
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   609
         _GAPHEIGHT      =   53
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Employee Code"
         Columns(0).DataField=   "employeecode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Employee Name"
         Columns(1).DataField=   "employeename"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=8361"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8281"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
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
         _PropDict       =   $"frmADOvertime2.frx":1AD0
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H404040&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
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
      Begin TDBText6Ctl.TDBText txtRemarks 
         Height          =   975
         Left            =   1680
         TabIndex        =   8
         Tag             =   "txtRegistrationRemarks"
         Top             =   975
         Width           =   6060
         _Version        =   65536
         _ExtentX        =   10689
         _ExtentY        =   1720
         Caption         =   "frmADOvertime2.frx":1B7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmADOvertime2.frx":1BE6
         Key             =   "frmADOvertime2.frx":1C04
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   765
         TabIndex        =   9
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hours"
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
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   270
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmADOvertime2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mAdd             As Boolean

Private Sub cmdClose_Click()

    Unload Me
    
    frmADOvertime.tdgOvertime.SetFocus
    
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrHndlr
    
    Dim rsDateTime              As ADODB.Recordset
    
    Dim mOtCode                 As Integer
    
    NetOpen rsDateTime, "select curdate() currentdate,curtime() currenttime"
    
    With frmADOvertime
    
        ConMain.Execute "set autocommit = 0"
        
        ConMain.BeginTrans
        
        If mAdd = True Then
        
            If Trim(tdbEmployee.Text) = "" Or IsNull(tdbEmployee.SelectedItem) Or tdbEmployee.ApproxCount <= 0 Then
                MsgBox "Please select an employee.", vbExclamation + vbOKOnly
                tdbEmployee.SetFocus
                Exit Sub
            End If
            
            If CDbl(txtOtHrs.Text) <= 0 Then
                MsgBox "Please enter a number greater than zero.", vbExclamation + vbOKOnly
                txtOtHrs.SetFocus
                Exit Sub
            End If
            
            mOtCode = LastCodeUsed("lastotcode", .mPerCode)
            ConMain.Execute "insert into overtimelne(otcode,employeecode,percode,othrs,status," & _
                            "tdatetime,fnlz,remarks) values " & _
                    "(" & mOtCode & "," & tdbEmployee.BoundText & ",'" & .mPerCode & "'," & Format(txtOtHrs.Text, "##0.00") & ",'Approved','" & _
                    Format(rsDateTime!currentdate, "YYYY-MM-DD") & "','N','" & Swap(txtRemarks.Text) & "')"
        Else
            If CDbl(txtOtHrs.Text) <= 0 Then
                MsgBox "Please enter a number greater than zero.", vbExclamation + vbOKOnly
                txtOtHrs.SetFocus
                Exit Sub
            End If
            
            mOtCode = .rsOvertime!otcode
            ConMain.Execute "update overtimelne set othrs = " & Format(txtOtHrs.Text, "##0.00") & ",remarks = '" & Swap(txtRemarks.Text) & "' where otcode = " & mOtCode & " and percode = " & .mPerCode & ""
            
        End If
        
        ConMain.Execute "update payroll set fnlz = 'N' where percode = " & .mPerCode & " and employeecode = " & tdbEmployee.BoundText & ""
        
        ConMain.Execute "update payrollperiod set genpay = 'N' where percode = " & .mPerCode & ""
        
        ConMain.CommitTrans
        
        Lock_Button "TTFTT", .cmdMenu, 4
        .rsOvertime.Requery
        .rsOvertime.MoveFirst
        .rsOvertime.Find "otcode = " & mOtCode & ""
        
        Unload Me
        
        .tdgOvertime.SetFocus
        
    End With
    
    Exit Sub
ErrHndlr:
    
    MsgBox "Error Message: " & err.Description, vbCritical + vbOKOnly
    
End Sub

Private Sub Form_Load()

    
    DoEvents
    If mAdd = False Then
        bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname,', ',firstname,' ',middlename) employeename from employee " & _
                        "where employeecode = " & frmADOvertime.rsOvertime!employeecode & " order by lastname,firstname,middlename", "employeename", "employeecode"
        tdbEmployee.Enabled = False
        tdbEmployee.BoundText = frmADOvertime.rsOvertime!employeecode
        txtOtHrs.Text = Format(frmADOvertime.rsOvertime!othrs, "#,##0.00")
        txtRemarks.Text = frmADOvertime.rsOvertime!remarks
    Else
        bind_tdb ConMain, tdbEmployee, "select employeecode,concat(lastname,', ',firstname,' ',middlename) employeename from employee where isactive <> 'N' order by lastname,firstname,middlename", "employeename", "employeecode"
    End If
    
End Sub

Private Sub tdbEmployee_GotFocus()
    With tdbEmployee
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tdbEmployee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbEmployee, tdbEmployee.RowSource, tdbEmployee.Text
    End If
End Sub

Private Sub txtOtHrs_GotFocus()
    With txtOtHrs
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtOtHrs_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRemarks_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
