VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmPPFinalizePayroll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finalize payroll"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   15
      TabIndex        =   0
      Top             =   -75
      Width           =   6105
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   105
         Left            =   2025
         TabIndex        =   1
         Top             =   660
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   105
         Left            =   2025
         TabIndex        =   2
         Top             =   795
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TrueOleDBList80.TDBCombo tdbPayrollPeriod 
         Height          =   345
         Left            =   1980
         TabIndex        =   3
         Tag             =   "Municipal"
         Top             =   180
         Width           =   4005
         _ExtentX        =   7064
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
         Columns(0).DataField=   "percode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descritpion"
         Columns(1).DataField=   "description"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "from"
         Columns(2).DataField=   "wrkdatefrom"
         Columns(2).NumberFormat=   "mm/dd/yyyy"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "To"
         Columns(3).DataField=   "wrkdateto"
         Columns(3).NumberFormat=   "mm/dd/yyyy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "payfreqcode"
         Columns(4).DataField=   "payfreqcode"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "payyear"
         Columns(5).DataField=   "payyear"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "paymonth"
         Columns(6).DataField=   "paymonth"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3254"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3016"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2937"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(37)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
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
         HeadLines       =   0
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
         _PropDict       =   $"frmPPFinalizePayroll.frx":0000
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF6F8F8&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   390
         Left            =   105
         TabIndex        =   4
         Top             =   585
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   688
         Caption         =   "&Generate"
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
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmPPFinalizePayroll.frx":00AA
         cBack           =   14737632
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   270
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPPFinalizePayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click()
    
    Dim rsPayroll               As ADODB.Recordset
    Dim rsSSSEC                 As ADODB.Recordset
    Dim rsLoans                 As ADODB.Recordset
    Dim rsLastLoanDed           As ADODB.Recordset
    Dim rsSSS                   As ADODB.Recordset
    Dim rsSSSTtlCont            As ADODB.Recordset
    
    Dim mLoanDedCode            As Integer

    Dim mBalance                As Double
    Dim mTtlAmtPaid             As Double
    Dim sssEc                   As Double
    
    If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then

        ConMain.Execute "set autocommit = 0"
        ConMain.BeginTrans

        NetOpen rsPayroll, "select x1.*,x2.payyear from lvavailed x1 " & _
                           "left outer join payrollperiod x2  on x1.percode=x2.percode " & _
                           "where x1.percode = " & tdbPayrollPeriod.BoundText & " order by x1.employeecode"
                           
        NetOpen rsLoans, "select * from loanded where percode = " & tdbPayrollPeriod.BoundText & " and " & _
                            "employeecode in (select employeecode from payroll where percode = " & tdbPayrollPeriod.BoundText & ") " & _
                            "order by loancode"
                            
'        NetOpen rsSSSEC, "select employeecode,payyear,paymonth,sum(sssamnt+ssser) sssttl from payroll " & _
'                         "where payyear = " & tdbPayrollPeriod.Columns("payyear").Value & " and paymonth = '" & tdbPayrollPeriod.Columns("paymonth").Value & "' " & _
'                         "group by employeecode,payyear,paymonth "
        
        NetOpen rsSSSTtlCont, "SELECT employeecode,count(employeecode) ctr,SUM(sssamnt+ssser) sssttl,sum(ec) ecttl " & _
                          "FROM payroll " & _
                          "WHERE payyear = " & tdbPayrollPeriod.Columns("payyear").Value & " and paymonth = '" & tdbPayrollPeriod.Columns("paymonth").Value & "' " & _
                          "group by employeecode "
                          
        'NetOpen rsLoans, "select * from loanded where percode = " & tdbPayrollPeriod.BoundText & " order by loancode"
        
        pb1.Max = rsPayroll.RecordCount + rsLoans.RecordCount + rsSSSTtlCont.RecordCount

        If pb1.Max > 0 Then

            With rsPayroll
                If .RecordCount > 0 Then

                    pb2.Value = 0
                    pb2.Max = .RecordCount

                    .MoveFirst

                    Do While Not .EOF

                        pb1.Value = pb1.Value + 1
                        pb2.Value = pb2.Value + 1
                        
                        ConMain.Execute "update lvlimit set lvlimit = " & !lvlimit & " " & _
                                        "where employeecode = " & !employeecode & " and " & _
                                        "payyear = " & !payYear & " and leavetypescode = " & !leavetypescode

                        .MoveNext
                      DoEvents
                    Loop

                End If
            End With

            With rsLoans

                If .RecordCount > 0 Then

                    pb2.Value = 0
                    pb2.Max = .RecordCount

                    ConMain.Execute "delete from loanded where percode = " & tdbPayrollPeriod.BoundText & ""

                    .MoveFirst

                    Do While Not .EOF

                        pb1.Value = pb1.Value + 1
                        pb2.Value = pb2.Value + 1

                        NetOpen rsLastLoanDed, "select ttlamtpaid,balance from loanded where loancode = " & !loancode & " order by loandedcode desc limit 1"

                        mBalance = rsLastLoanDed!balance - !amtded
                        mTtlAmtPaid = rsLastLoanDed!ttlamtpaid + !amtded

                        mLoanDedCode = LastLoanCodeUsed(!loancode)

                        ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,percode," & _
                                        "payyear,paymonth,amtded,dateposted,ttlamtpaid, " & _
                                        "balance,fnlz,cancelled,remarks,usercode) values " & _
                                      "(" & mLoanDedCode & "," & !loancode & "," & !loantypescode & "," & !employeecode & "," & !percode & "," & _
                                       !payYear & ",'" & !payMonth & "'," & !amtded & ",'" & Format(!dateposted, "YYYY-MM-DD") & "', " & mTtlAmtPaid & " ," & _
                                       mBalance & ",'Y','N',''," & !UserCode & ")"

                        If CDbl(mBalance) = 0 Then
                            ConMain.Execute "update loans set status = 'Paid' where loancode = " & CInt(!loancode) & ""
                        ElseIf CDbl(mBalance) < 0 Then
                            ConMain.Execute "update loans set status = 'Over Paid' where loancode = " & CInt(!loancode) & ""
                        End If

                        .MoveNext
                      DoEvents
                    Loop

                End If

            End With
            
'            With rsSSSEC
'              If .RecordCount > 0 Then
'
'                pb2.Value = 0
'                pb2.Max = .RecordCount
'
'                ConMain.Execute "delete from sss_monthly_ec where payyear = " & tdbPayrollPeriod.Columns("payyear").Value & " and paymonth='" & tdbPayrollPeriod.Columns("paymonth").Value & "' "
'                NetOpen rsSSS, "select er+ee contttl,ec from sss order by SSSBCode"
'                .MoveFirst
'                Do While Not .EOF
'
'                    pb1.Value = pb1.Value + 1
'                    pb2.Value = pb2.Value + 1
'
'                  rsSSS.MoveFirst
'                  Do While Not rsSSS.EOF
'                      If rsSSS.AbsolutePosition = rsSSS.RecordCount Then
'                        ConMain.Execute "insert into sss_monthly_ec (employeecode,payyear,paymonth,ec) values  " & _
'                                        "(" & !employeecode & "," & !payyear & ",'" & !paymonth & "'," & rsSSS!ec & ")"
'                        Exit Do
'                      End If
'                      If CDbl(rsSSS!contttl) = CDbl(!sssttl) Then
'                        ConMain.Execute "insert into sss_monthly_ec (employeecode,payyear,paymonth,ec) values  " & _
'                                      "(" & !employeecode & "," & !payyear & ",'" & !paymonth & "'," & rsSSS!ec & ")"
'                        Exit Do
'                      End If
'
'                    rsSSS.MoveNext
'                  Loop
'
'                  .MoveNext
'                Loop
'              End If
'            End With
            
'            With rsSSSTtlCont
'              If .RecordCount > 0 Then
'                ConMain.Execute "delete from sss_monthly_ec where payyear = " & tdbPayrollPeriod.Columns("payyear").Value & " and paymonth='" & tdbPayrollPeriod.Columns("paymonth").Value & "' "
'                .MoveFirst
'                pb2.Value = 0
'                pb2.Max = .RecordCount
'                Do While Not .EOF
'                  pb1.Value = pb1.Value + 1
'                  pb2.Value = pb2.Value + 1
'
'                  sssEc = 0
'
'                  If CDbl(!sssttl) = 0 Then
'                    sssEc = 0
'                  ElseIf CDbl(!sssttl) > 0 And CDbl(!sssttl) < 1650 Then
'                    ConMain.Execute "insert into sss_monthly_ec (employeecode,payyear,paymonth,ec) values  " & _
'                                      "(" & !employeecode & "," & tdbPayrollPeriod.Columns("payyear").Value & _
'                                      ",'" & tdbPayrollPeriod.Columns("paymonth").Value & "',10)"
'                    If CDbl(!ecttl) <> 10 Then sssEc = 10 / !CTR
'                  ElseIf CDbl(!sssttl) >= 1650 Then
'                    ConMain.Execute "insert into sss_monthly_ec (employeecode,payyear,paymonth,ec) values  " & _
'                                      "(" & !employeecode & "," & tdbPayrollPeriod.Columns("payyear").Value & _
'                                      ",'" & tdbPayrollPeriod.Columns("paymonth").Value & "',30)"
'                    If CDbl(!ecttl) <> 30 Then sssEc = 30 / !CTR
'                  End If
'
'                  If CDbl(!CTR > 1) Then
'                    If sssEc > 0 Then
'                      ConMain.Execute "update payroll set ec = " & sssEc & " where employeecode = " & !employeecode & " and " & _
'                                    "payyear = " & tdbPayrollPeriod.Columns("payyear").Value & " and " & _
'                                    "paymonth = '" & tdbPayrollPeriod.Columns("paymonth").Value & "' "
'                    End If
'                  End If
'
'                  .MoveNext
'                  DoEvents
'                Loop
'              End If
'            End With

        End If
        
        ConMain.Execute "update payroll_sss_contributions set fnlz = 'Y' where percode = " & tdbPayrollPeriod.BoundText & ""
                           
        ConMain.Execute "update payrollperiod set fnlz = 'Y' where percode = " & tdbPayrollPeriod.BoundText & ""

        ConMain.Execute "update payroll set fnlz = 'Y' where percode = " & tdbPayrollPeriod.BoundText & ""

        ConMain.CommitTrans

        MsgBox "Process completed.", vbInformation + vbOKOnly

        tdbPayrollPeriod.SetFocus

    Else

        MsgBox "Please select a payroll period.", vbExclamation + vbOKOnly
        tdbPayrollPeriod.SetFocus

    End If
    'RecalLoans
End Sub

Private Sub tdbPayrollPeriod_GotFocus()

    If Trim(tdbPayrollPeriod.Text) <> "" And Not IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount > 0 Then
        tdbPayrollPeriod.Tag = tdbPayrollPeriod.BoundText
    Else
        tdbPayrollPeriod.Tag = ""
    End If
        
    bind_tdb ConMain, tdbPayrollPeriod, "select percode,description,wrkdatefrom,wrkdateto,payfreqcode,payyear,paymonth from payrollperiod where fnlz = 'N' and genpay = 'Y' order by percode desc", "description", "percode"


    tdbPayrollPeriod.BoundText = tdbPayrollPeriod.Tag
    
    If IsNull(tdbPayrollPeriod.SelectedItem) And tdbPayrollPeriod.ApproxCount <= 0 Then
        tdbPayrollPeriod.BoundText = ""
    End If
    
    tdbPayrollPeriod.Tag = ""
    
    tdbPayrollPeriod.SelStart = 0
    tdbPayrollPeriod.SelLength = Len(tdbPayrollPeriod.Text)
    
    
End Sub

Private Sub tdbPayrollPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        SearchList KeyAscii, tdbPayrollPeriod, tdbPayrollPeriod.RowSource, tdbPayrollPeriod.Text
    End If
End Sub

Private Sub RecalLoans()

  Dim rsLoans               As ADODB.Recordset
  Dim rsLastLoanDed         As ADODB.Recordset
  
  Dim mTtlAmtPaid           As Double
  Dim mBalance              As Double
  
  NetOpen rsLoans, "select * from loanded where percode = 14 and " & _
                            "employeecode in (select employeecode from payroll where percode = 14) " & _
                            "order by loancode"
  
  With rsLoans
    If .RecordCount Then
    
      ConMain.Execute "set autocommit = 0"
      ConMain.BeginTrans
      
      .MoveFirst
      
      Do While Not .EOF
      
        NetOpen rsLastLoanDed, "select loandedcode from loanded where loancode = " & !loancode & " and percode < 14 order by loandedcode desc limit 1"
        If rsLastLoanDed.RecordCount > 0 Then
          ConMain.Execute "update loanded set loandedcode = " & rsLastLoanDed!loandedcode + 1 & " where loancode = " & !loancode & " and percode =14"
        Else
          ConMain.Execute "update loanded set loandedcode = 1 where loancode = " & !loancode & " and percode =14"
        End If
        
        ConMain.Execute "update loanded set fnlz = 'Y' where percode = 14 and loancode = " & !loancode & ""
        
        If !balance > 0 Then
          
          NetOpen rsLastLoanDed, "select loandedcode,percode,amtded from loanded where loancode = " & !loancode & " and percode > 14 order by percode"
          If rsLastLoanDed.RecordCount > 0 Then
            
            mTtlAmtPaid = !ttlamtpaid
            mBalance = !balance
            
            rsLastLoanDed.MoveFirst
            
            Do While Not rsLastLoanDed.EOF
            
              mTtlAmtPaid = mTtlAmtPaid + rsLastLoanDed!amtded
              mBalance = mBalance - !amtded
              
              ConMain.Execute "update loanded set ttlamtpaid = " & mTtlAmtPaid & ", balance = " & mBalance & " " & _
                              "where percode = " & rsLastLoanDed!percode & " and loancode = " & !loancode & ""
                                          
              If CDbl(mBalance) = 0 Then
                  ConMain.Execute "update loans set status = 'Paid' where loancode = " & CInt(!loancode) & ""
              ElseIf CDbl(mBalance) < 0 Then
                  ConMain.Execute "update loans set status = 'Over Paid' where loancode = " & CInt(!loancode) & ""
              End If
              
              rsLastLoanDed.MoveNext
              
            Loop
              
          End If
        
        ElseIf !balance = 0 Then
          ConMain.Execute "update loans set status = 'Paid' where loancode = " & CInt(!loancode) & ""
        ElseIf !balance < 0 Then
          ConMain.Execute "update loans set status = 'Over Paid' where loancode = " & CInt(!loancode) & ""
        End If
        
        .MoveNext
        
      Loop
      
      ConMain.CommitTrans
    End If
  End With
  MsgBox "Done!"
End Sub
