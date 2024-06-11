VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtilImportAll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Utility"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImportLvLimit 
      Caption         =   "&Import"
      Height          =   375
      Left            =   7785
      TabIndex        =   9
      Top             =   6015
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "&Save As..."
      Height          =   375
      Left            =   1815
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdImportLoans 
      Caption         =   "&Import"
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cmdImportOtherdedFinal 
      Caption         =   "&Import"
      Height          =   375
      Left            =   7815
      TabIndex        =   6
      Top             =   2235
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdImportOtherDed 
      Caption         =   "&Import"
      Height          =   375
      Left            =   8100
      TabIndex        =   5
      Top             =   3105
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnPreviousSheet 
      Caption         =   "&Previous Sheet"
      Enabled         =   0   'False
      Height          =   372
      Left            =   3465
      TabIndex        =   4
      Top             =   60
      Width           =   1692
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Import"
      Height          =   375
      Left            =   8325
      TabIndex        =   2
      Top             =   3735
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnNextSheet 
      Caption         =   "&Next Sheet"
      Enabled         =   0   'False
      Height          =   372
      Left            =   5175
      TabIndex        =   1
      Top             =   60
      Width           =   1692
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7890
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   6840
      _cx             =   12065
      _cy             =   13917
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7485
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.xls"
   End
End
Attribute VB_Name = "frmUtilImportAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const APPNAME = "VSFlexGrid8 <-> Excel"
Dim sheet%

Private Sub btnExport_Click()
        
    On Error GoTo ErrHndlr

    dlg.FileName = ""
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub

    MousePointer = MousePointerConstants.vbHourglass
    fg.SaveGrid dlg.FileName, flexFileExcel
    MousePointer = MousePointerConstants.vbDefault
    
    Caption = APPNAME + " [" + dlg.FileName + "]"

    sheet = 0
    btnNextSheet.Enabled = False
    
ErrHndlr:

End Sub

Private Sub cmdImportLoans_Click()

    Dim mBranchCode             As String
    Dim mDivisionCode           As String
    Dim mCostCenterCode         As String
    Dim mSectionCode            As String
    
    Dim mLoanCode               As Integer
    Dim mLoanDedCode            As Integer
    
    Dim R                       As Long
    
    Dim rsEmpCheck              As ADODB.Recordset
    
    Dim mWithNoError            As Boolean
    
    NetOpen rsEmpCheck, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode," & _
                        "concat(lastname,', ',firstname) employeename,concat(lastname,', ',firstname,' ',middlename) employeename2 from employee order by lastname,firstname "
    
    If rsEmpCheck.RecordCount > 0 Then
        
        With fg
            If .Rows > 0 Then
                
                R = 0
                
                mWithNoError = True
                
                '.Cols = 2
                For R = 0 To .Rows - 1
                
                    If Not IsNumeric(.TextMatrix(R, 2)) Then
                        .Cell(flexcpBackColor, R, 2, R, 2) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 2, R, 2) = vbWhite
                    End If
                    
                    If Not IsNumeric(.TextMatrix(R, 3)) Then
                        .Cell(flexcpBackColor, R, 3, R, 3) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 3, R, 3) = vbWhite
                    End If
                    
                    If Not IsNumeric(.TextMatrix(R, 4)) Then
                        .Cell(flexcpBackColor, R, 4, R, 4) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 4, R, 4) = vbWhite
                    End If
                    
                    If Not IsNumeric(.TextMatrix(R, 5)) Then
                        .Cell(flexcpBackColor, R, 5, R, 5) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 5, R, 5) = vbWhite
                    End If
                    
                    rsEmpCheck.MoveFirst
                    rsEmpCheck.Find "employeename = '" & Trim(.TextMatrix(R, 0)) & ", " & Trim(.TextMatrix(R, 1)) & "'"
                    
                    If rsEmpCheck.EOF Then
                        rsEmpCheck.MoveFirst
                        rsEmpCheck.Find "employeename2 = '" & Trim(.TextMatrix(R, 0)) & ", " & Trim(.TextMatrix(R, 1)) & "'"
                        If rsEmpCheck.EOF Then
                            .Cell(flexcpBackColor, R, 0, R, 1) = vbRed
                            mWithNoError = False
                        Else
                            .Cell(flexcpBackColor, R, 0, R, 1) = vbWhite
                        End If
                    Else
                        .Cell(flexcpBackColor, R, 0, R, 1) = vbWhite
                    End If
                    
                Next
                
                If Not mWithNoError Then
                    MsgBox "One or more rows contain employee name that doesn't found in the database or amount with invalid numbers.", vbExclamation + vbOKOnly
                    Exit Sub
                End If
                
                R = 0
                
                ConMain.Execute "set autocommit = 0;"
                ConMain.BeginTrans
                
                For R = 0 To .Rows - 1
                
                        rsEmpCheck.MoveFirst
                        rsEmpCheck.Find "employeename = '" & Trim(.TextMatrix(R, 0)) & ", " & Trim(.TextMatrix(R, 1)) & "'"
                        If rsEmpCheck.EOF Then
                            rsEmpCheck.MoveFirst
                            rsEmpCheck.Find "employeename2 = '" & Trim(.TextMatrix(R, 0)) & ", " & Trim(.TextMatrix(R, 1)) & "'"
                        End If
                        
                        If Not IsNumeric(rsEmpCheck!branchcode) Or IsNull(rsEmpCheck!branchcode) Then
                            mBranchCode = "Null"
                        Else
                            mBranchCode = rsEmpCheck!branchcode
                        End If
                        
                        If Not IsNumeric(rsEmpCheck!divisioncode) Or IsNull(rsEmpCheck!divisioncode) Then
                            mDivisionCode = "Null"
                        Else
                            mDivisionCode = rsEmpCheck!divisioncode
                        End If
                        
                        If Not IsNumeric(rsEmpCheck!costcentercode) Or IsNull(rsEmpCheck!costcentercode) Then
                            mCostCenterCode = "Null"
                        Else
                            mCostCenterCode = rsEmpCheck!costcentercode
                        End If
                        
                        If Not IsNumeric(rsEmpCheck!sectioncode) Or IsNull(rsEmpCheck!sectioncode) Then
                            mSectionCode = "Null"
                        Else
                            mSectionCode = rsEmpCheck!sectioncode
                        End If
                        
                        mLoanCode = LastCode("Loans")
                
                        ConMain.Execute "insert into loans(loancode,dummycode,employeecode,loantypescode,costcentercode," & _
                                "divisioncode,branchcode,loandate,loanamnt,dedperpayday, " & _
                                "noofinst,startdate,status,remarks,referenceno) values (" & _
                                mLoanCode & ",'" & Format(mLoanCode, "0000000000") & "', " & rsEmpCheck!employeecode & ",1," & mCostCenterCode & ", " & _
                                mDivisionCode & "," & mBranchCode & ",'2009-05-10'," & Format(.TextMatrix(R, 2), "##0.00") & "," & Format(.TextMatrix(R, 5), "##0.00") & ", " & _
                                Format(.TextMatrix(R, 6), "##0") & ",'2009-05-10'," & "'Active','','')"

                        mLoanDedCode = LastLoanCodeUsed(mLoanCode)
                
                        ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled) values " & _
                                "(" & mLoanDedCode & "," & mLoanCode & ",1," & rsEmpCheck!employeecode & "," & _
                                 Format(.TextMatrix(R, 4), "##0.00") & ",'2009-05-12', " & Format(.TextMatrix(R, 4), "##0.00") & "," & Format(.TextMatrix(R, 3), "##0.00") & ",'Y','N')"
                Next
                
                ConMain.CommitTrans
                MsgBox "Data was succesfully imported.", vbInformation + vbOKOnly
                
            End If
        End With
        
    End If
End Sub

Private Sub cmdImportLvLimit_Click()

    Dim rsEmpCheck              As ADODB.Recordset
    
    Dim mWithNoError            As Boolean
    
    Dim R                       As Long
    
    NetOpen rsEmpCheck, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode from employee order by lastname,firstname "
    
    If rsEmpCheck.RecordCount > 0 Then
        With fg
            If .Rows > 0 Then
                
                
                R = 0
                
                mWithNoError = True
                
                '.Cols = 2
                For R = 0 To .Rows - 1
                
                    If Not IsNumeric(.TextMatrix(R, 1)) Then
                        .Cell(flexcpBackColor, R, 1, R, 1) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 1, R, 1) = vbWhite
                    End If
                    
                    If Not IsNumeric(.TextMatrix(R, 2)) Then
                        .Cell(flexcpBackColor, R, 2, R, 2) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 2, R, 2) = vbWhite
                    End If
                    
                    
                    rsEmpCheck.MoveFirst
                    rsEmpCheck.Find "employeecode = '" & Trim(.TextMatrix(R, 0)) & "'"
                    
                    If rsEmpCheck.EOF Then
                        .Cell(flexcpBackColor, R, 0, R, 1) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 0, R, 1) = vbWhite
                    End If
                    
                Next
                
                If Not mWithNoError Then
                    MsgBox "One or more rows contain employee name that doesn't found in the database or amount with invalid numbers.", vbExclamation + vbOKOnly
                    Exit Sub
                End If
                
                R = 0
                
                ConMain.Execute "set autocommit = 0;"
                ConMain.BeginTrans
                
                For R = 0 To .Rows - 1
                
                        ConMain.Execute "insert into lvlimit(payyear,employeecode,leavetypescode,lvlimit) values (" & _
                                        2009 & ", " & .TextMatrix(R, 0) & ",1," & Format(.TextMatrix(R, 1), "##0.00") & ") "
                        
                        ConMain.Execute "insert into lvlimit(payyear,employeecode,leavetypescode,lvlimit) values (" & _
                                        2009 & ", " & .TextMatrix(R, 0) & ",3," & Format(.TextMatrix(R, 2), "##0.00") & ") "
                                        
                Next
                
                ConMain.CommitTrans
                MsgBox "Data was succesfully imported.", vbInformation + vbOKOnly
                
            End If
        End With
    End If
    
End Sub

Private Sub cmdImportOtherDed_Click()
    Dim R               As Long
    
    With fg
        If .Rows > 0 Then
            ConMain.Execute "set autocommit = 0;"
            ConMain.BeginTrans
            ConMain.Execute "delete from importotherded where otherdeductionscode = 4"
            For R = 0 To fg.Rows - 1
                   ConMain.Execute "insert into importotherded (fullname,amount,otherdeductionscode) values ('" & Swap(.TextMatrix(R, 0)) & "'," & Format(.TextMatrix(R, 1), "##0.00") & ",4)"
            Next
            ConMain.CommitTrans
            MsgBox "Data was succesfully imported!", vbInformation + vbOKOnly
        End If
    End With
End Sub

Private Sub cmdImportOtherdedFinal_Click()

    Dim rs                  As ADODB.Recordset
    Dim mBranchCode         As String
    Dim mDivisionCode       As String
    Dim mCostCenterCode     As String
    Dim mSectionCode        As String

    NetOpen rs, "select x1.*,x2.branchcode,x2.divisioncode,x2.costcentercode,x2.sectioncode," & _
                "x3.otherdeductionsname,x2.employeecode from importotherded x1 " & _
                "left outer join employee x2 on x1.fullname =  concat(x2.lastname,', ', x2.firstname) " & _
                "left outer join otherdeductions x3 on x1.otherdeductionscode = x3.otherdeductionscode " & _
                "Where X2.employeecode Is Not Null " & _
                "order by x2.employeecode,x1.fullname "
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            ConMain.Execute "set autocommit = 0;"
            ConMain.BeginTrans
            Do While Not .EOF
            
                If Not IsNumeric(!branchcode) Or IsNull(!branchcode) Then
                    mBranchCode = "Null"
                Else
                    mBranchCode = !branchcode
                End If
                
                If Not IsNumeric(!divisioncode) Or IsNull(!divisioncode) Then
                    mDivisionCode = "Null"
                Else
                    mDivisionCode = !divisioncode
                End If
                
                If Not IsNumeric(!costcentercode) Or IsNull(!costcentercode) Then
                    mCostCenterCode = "Null"
                Else
                    mCostCenterCode = !costcentercode
                End If
                
                If Not IsNumeric(!sectioncode) Or IsNull(!sectioncode) Then
                    mSectionCode = "Null"
                Else
                    mSectionCode = !sectioncode
                End If
                
                ConMain.Execute "insert into deductions (otherdeductionscode,percode,employeecode,costcentercode,divisioncode," & _
                        "branchcode,sectioncode,payyear,paymonth,amount,remarks) values ( " & _
                        "" & !OtherDeductionscode & "," & 2 & "," & !employeecode & ", " & mCostCenterCode & "," & mDivisionCode & ", " & _
                        "" & mBranchCode & "," & mSectionCode & ", '2009','April'," & !amount & ",'')"
                .MoveNext
                
            Loop
            ConMain.CommitTrans
            MsgBox "Data was imported succesfully!", vbInformation + vbOKOnly
        End If
    End With
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        fg.RemoveItem fg.Row
    Else
    
    End If
End Sub

Private Sub Form_Load()

    'Caption = APPNAME
    
    dlg.DefaultExt = "xls"
    dlg.Filter = "Excel 97 (*.xls)|*.xls"
    
    fg.AllowUserResizing = flexResizeBoth
    fg.MergeCells = flexMergeSpill
    fg.ExtendLastCol = True
    
End Sub

Private Sub Form_Resize()

'    On Error Resume Next
'    With fg
'    .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left
'    End With
    
End Sub

Private Sub btnLoad_Click()

    dlg.FileName = ""
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub

    MousePointer = MousePointerConstants.vbHourglass
    fg.LoadGrid dlg.FileName, flexFileExcel
    MousePointer = MousePointerConstants.vbDefault
    
    Caption = APPNAME + " [" + dlg.FileName + "]"

    sheet = 0
    btnNextSheet.Enabled = True
    btnPreviousSheet.Enabled = True
    
    

End Sub

Private Sub btnSave_Click()

'    dlg.FileName = ""
'    dlg.ShowOpen
'    If Len(dlg.FileName) = 0 Then Exit Sub
'
'    MousePointer = MousePointerConstants.vbHourglass
'    fg.SaveGrid dlg.FileName, flexFileExcel
'    MousePointer = MousePointerConstants.vbDefault
'
'    Caption = APPNAME + " [" + dlg.FileName + "]"
'
'    sheet = 0
'    btnNextSheet.Enabled = False
'    btnPreviousSheet.Enabled = False
    Dim R               As Long
    
    With fg
        If .Rows > 0 Then
            ConMain.Execute "set autocommit = 0;"
            ConMain.BeginTrans
            ConMain.Execute "delete from acctno"
            For R = 0 To fg.Rows - 1
                   ConMain.Execute "insert into acctno (acctno,lname,fname) values (" & .TextMatrix(R, 0) & ",'" & Swap(.TextMatrix(R, 1)) & "','" & Swap(.TextMatrix(R, 2)) & "')"
            Next
            ConMain.CommitTrans
            MsgBox "Data was succesfully imported!", vbInformation + vbOKOnly
        End If
    End With

End Sub

Private Sub btnNextSheet_Click()

    sheet = sheet + 1
    
    MousePointer = MousePointerConstants.vbHourglass
    
    On Error Resume Next
    fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    If err <> 0 Then
        sheet = sheet - 1
        fg.LoadGrid dlg.FileName, flexFileExcel, sheet
'        MsgBox "No More Sheets"
'        sheet = 0
'        btnNextSheet.Enabled = False
    End If
    On Error GoTo 0
    
    MousePointer = MousePointerConstants.vbDefault

    
End Sub

Private Sub btnPreviousSheet_Click()

    sheet = sheet - 1
    
    MousePointer = MousePointerConstants.vbHourglass
    
    On Error Resume Next
    fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    If err <> 0 Then
        sheet = sheet + 1
        fg.LoadGrid dlg.FileName, flexFileExcel, sheet
'        MsgBox "No More Sheets"
'        sheet = 0
'        btnNextSheet.Enabled = False
    End If
    On Error GoTo 0
    
    MousePointer = MousePointerConstants.vbDefault

    
End Sub





