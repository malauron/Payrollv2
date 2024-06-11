VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilImportOtherded 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7920
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   8820
      _cx             =   15557
      _cy             =   13970
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
      AllowUserResizing=   1
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
      Editable        =   2
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
      Left            =   2835
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.xls"
   End
   Begin lvButton.lvButtons_H btnLoad 
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   45
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Load Excel File"
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
      Image           =   "frmUtilImportOtherded.frx":0000
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnPreviousSheet 
      Height          =   375
      Left            =   5430
      TabIndex        =   3
      Top             =   45
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Previous Sheet"
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
      Image           =   "frmUtilImportOtherded.frx":077A
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnNextSheet 
      Height          =   375
      Left            =   7125
      TabIndex        =   4
      Top             =   45
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Next Sheet"
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
      ImgAlign        =   2
      Image           =   "frmUtilImportOtherded.frx":0EF4
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H cmdImportOtherded 
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   8415
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
      Caption         =   "&Import to Database"
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
      Image           =   "frmUtilImportOtherded.frx":166E
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnSave 
      Height          =   375
      Left            =   6540
      TabIndex        =   6
      Top             =   8415
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
      Caption         =   "&Save as Excel File"
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
      Image           =   "frmUtilImportOtherded.frx":2348
      cBack           =   14737632
   End
   Begin VB.Label lblOthers 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Type: "
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
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   8475
      Width           =   4110
   End
End
Attribute VB_Name = "frmUtilImportOtherded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mImportType          As String
Public mVoucherType         As String
Public mDeductionType       As String

Public mTypeCode            As Integer

Const APPNAME = "VSFlexGrid8 <-> Excel"
Dim sheet%

Private Sub cmdImportOtherDed_Click()

    Dim mBranchCode             As String
    Dim mDivisionCode           As String
    Dim mCostCenterCode         As String
    Dim mSectionCode            As String

    Dim R                       As Long
    
    Dim mTransCode              As Integer
    
    Dim mTotal                  As Double
    
    Dim rsEmpCheck              As ADODB.Recordset
    
    Dim mWithNoError            As Boolean
    
    
    NetOpen rsEmpCheck, "select employeecode,branchcode,divisioncode,costcentercode,sectioncode," & _
                        "concat(lastname,', ',firstname) employeename,concat(lastname,', ',firstname,' ',middlename) employeename2 from employee where isactive = 'Y' order by lastname,firstname "
    
    If rsEmpCheck.RecordCount > 0 Then
        With fg
            If .Rows > 0 Then
                
                R = 0
                
                mWithNoError = True
                
                .Cols = 2
                For R = 0 To .Rows - 1
                
                    If Not IsNumeric(.TextMatrix(R, 1)) Then
                        .Cell(flexcpBackColor, R, 1, R, 1) = vbRed
                        mWithNoError = False
                    Else
                        .Cell(flexcpBackColor, R, 1, R, 1) = vbWhite
                    End If
                    
                    rsEmpCheck.MoveFirst
                    rsEmpCheck.Find "employeename = '" & Trim(.TextMatrix(R, 0)) & "'"
                    
                    If rsEmpCheck.EOF Then
                        rsEmpCheck.MoveFirst
                        rsEmpCheck.Find "employeename2 = '" & Trim(.TextMatrix(R, 0)) & "'"
                        If rsEmpCheck.EOF Then
                            .Cell(flexcpBackColor, R, 0, R, 0) = vbRed
                            mWithNoError = False
                        Else
                            .Cell(flexcpBackColor, R, 0, R, 0) = vbWhite
                        End If
                    Else
                        .Cell(flexcpBackColor, R, 0, R, 0) = vbWhite
                    End If
                    
                Next
                
                If Not mWithNoError Then
                    MsgBox "One or more rows contain employee name that is not found in the database or amount with invalid numbers.", vbExclamation + vbOKOnly
                    Exit Sub
                End If
                
                If MsgBox("Do you want to import this to database?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                
                R = 0
                
                ConMain.Execute "set autocommit = 0;"
                ConMain.BeginTrans
                
'                If mImportType = "Otherded" Then
'                    ConMain.Execute "delete from deductions where otherdeductionscode = " & mTypeCode & " and percode = " & frmAdOtherDeductions.mPerCode & ""
'                Else
'                    ConMain.Execute "delete from earnings where otherearningscode = " & mTypeCode & " and percode = " & frmAdOtherEarnings.mPerCode & ""
'                End If
                
                For R = 0 To .Rows - 1
                
                        rsEmpCheck.MoveFirst
                        rsEmpCheck.Find "employeename = '" & Trim(.TextMatrix(R, 0)) & "'"
                        If rsEmpCheck.EOF Then
                            rsEmpCheck.MoveFirst
                            rsEmpCheck.Find "employeename2 = '" & Trim(.TextMatrix(R, 0)) & "'"
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
                        
                        If mImportType = "Otherded" Then
                            mTransCode = LastCodeUsed("lastotherdeductionslnecode", frmAdOtherDeductions.mPerCode)
                            ConMain.Execute "insert into deductions (otherdeductionscode,percode,employeecode,costcentercode,divisioncode," & _
                                    "branchcode,sectioncode,payyear,paymonth,amount," & _
                                    "remarks,otherdeductionslnecode) values (" & _
                                    mTypeCode & "," & frmAdOtherDeductions.mPerCode & "," & rsEmpCheck!employeecode & ", " & mCostCenterCode & "," & mDivisionCode & "," & _
                                    mBranchCode & "," & mSectionCode & ", '" & frmAdOtherDeductions.tdbPayrollPeriod.Columns("payyear").Text & "','" & frmAdOtherDeductions.tdbPayrollPeriod.Columns("paymonth").Text & "'," & Format(.TextMatrix(R, 1), "##0.00") & "," & _
                                    "''," & mTransCode & ")"
                        ElseIf mImportType = "OtherEarnings" Then
                            mTransCode = LastCodeUsed("lastotherearningslnecode", frmAdOtherEarnings.mPerCode)
                            ConMain.Execute "insert into earnings (otherearningscode,percode,employeecode,costcentercode,divisioncode," & _
                                    "branchcode,sectioncode,payyear,paymonth,amount," & _
                                    "remarks,otherearningslnecode) values ( " & _
                                    mTypeCode & "," & frmAdOtherEarnings.mPerCode & "," & rsEmpCheck!employeecode & ", " & mCostCenterCode & "," & mDivisionCode & ", " & _
                                    mBranchCode & "," & mSectionCode & ", '" & frmAdOtherEarnings.tdbPayrollPeriod.Columns("payyear").Text & "','" & frmAdOtherEarnings.tdbPayrollPeriod.Columns("paymonth").Text & "'," & Format(.TextMatrix(R, 1), "##0.00") & "," & _
                                    "''," & mTransCode & ")"
                        ElseIf mImportType = "Overtime" Then
                            mTransCode = LastCodeUsed("lastotcode", frmADOvertime.mPerCode)
                            ConMain.Execute "insert into overtimelne(otcode,employeecode,percode,othrs,status," & _
                                    "tdatetime,fnlz,remarks) values " & _
                                    "(" & mTransCode & "," & rsEmpCheck!employeecode & ",'" & frmADOvertime.mPerCode & "'," & Format(.TextMatrix(R, 1), "##0.00") & ",'Approved',DATE(NOW())," & _
                                    "'N','')"
                        ElseIf mImportType = "Vouchers" Then
                            mTransCode = LastCode(frmAdVouchers.mVoucherType & " - VoucherCode")
                            ConMain.Execute "insert into vouchers (vouchercode,dummycode,employeecode,amount,dateissued, " & _
                                    "vouchervalidperiodcode,vouchertype," & _
                                    "branchcode,divisioncode,costcentercode,sectioncode) values ( " & _
                                     mTransCode & ",'" & Format(mTransCode, "00000000000") & "', " & rsEmpCheck!employeecode & "," & Format(.TextMatrix(R, 1), "##0.00") & ",curdate()," & _
                                     frmAdVouchers.mVoucherValidPeriodCode & ",'" & frmAdVouchers.mVoucherType & "'," & _
                                     mBranchCode & "," & mDivisionCode & "," & mCostCenterCode & "," & mSectionCode & ")"
                        End If
                        
                Next
                
                ConMain.CommitTrans
                
                MsgBox "Data was succesfully imported!", vbInformation + vbOKOnly
                
                Unload Me
                
                If mImportType = "Otherded" Then
                
                    frmAdOtherDeductions.rsOtherDeductions.Requery
                    
                    If frmAdOtherDeductions.rsOtherDeductions.RecordCount > 0 Then
                    
                        frmAdOtherDeductions.txtNoOfRecords.Text = Format(frmAdOtherDeductions.rsOtherDeductions.RecordCount, "#,##0")
                        
                        Lock_Button "TTFTTT", frmAdOtherDeductions.cmdMenu, 5
                        
                        frmAdOtherDeductions.rsOtherDeductions.MoveFirst
                        
                        Do While Not frmAdOtherDeductions.rsOtherDeductions.EOF
                            mTotal = mTotal + frmAdOtherDeductions.rsOtherDeductions!amount
                            frmAdOtherDeductions.rsOtherDeductions.MoveNext
                        Loop
                        
                        frmAdOtherDeductions.rsOtherDeductions.MoveFirst
                        frmAdOtherDeductions.txtTotal.Text = Format(mTotal, "#,##0.00")
                        
                    Else
                        
                        frmAdOtherDeductions.txtNoOfRecords.Text = Format(frmAdOtherDeductions.rsOtherDeductions.RecordCount, "#,##0")
                        Lock_Button "TFFFTT", frmAdOtherDeductions.cmdMenu, 5
                        frmAdOtherDeductions.txtTotal.Text = "0.00"
                        
                    End If
                    
                ElseIf mImportType = "OtherEarnings" Then
                
                    frmAdOtherEarnings.rsOtherEarnings.Requery
                    
                    If frmAdOtherEarnings.rsOtherEarnings.RecordCount > 0 Then
                    
                        frmAdOtherEarnings.txtNoOfRecords.Text = Format(frmAdOtherEarnings.rsOtherEarnings.RecordCount, "#,##0")
                        
                        Lock_Button "TTFTTT", frmAdOtherEarnings.cmdMenu, 5
                        
                        frmAdOtherEarnings.rsOtherEarnings.MoveFirst
                        
                        Do While Not frmAdOtherEarnings.rsOtherEarnings.EOF
                            mTotal = mTotal + frmAdOtherEarnings.rsOtherEarnings!amount
                            frmAdOtherEarnings.rsOtherEarnings.MoveNext
                        Loop
                        
                        frmAdOtherEarnings.rsOtherEarnings.MoveFirst
                        frmAdOtherEarnings.txtTotal.Text = Format(mTotal, "#,##0.00")
                        
                    Else
                        
                        frmAdOtherEarnings.txtNoOfRecords.Text = Format(frmAdOtherEarnings.rsOtherEarnings.RecordCount, "#,##0")
                        Lock_Button "TFFFTT", frmAdOtherEarnings.cmdMenu, 5
                        frmAdOtherEarnings.txtTotal.Text = "0.00"
                        
                    End If
                ElseIf mImportType = "Overtime" Then
                    frmADOvertime.rsOvertime.Requery
                    If frmADOvertime.rsOvertime.RecordCount > 0 Then
                      frmADOvertime.rsOvertime.MoveFirst
                    End If
                    
                ElseIf mImportType = "Vouchers" Then
                
                    frmAdVouchers.rsVouchers.Requery
                    
                    If frmAdVouchers.rsVouchers.RecordCount > 0 Then
                    
                        frmAdVouchers.txtNoOfRecords.Text = Format(frmAdVouchers.rsVouchers.RecordCount, "#,##0")
                        
                        Lock_Button "TTFTTT", frmAdVouchers.cmdMenu, 5
                        
                        frmAdVouchers.rsVouchers.MoveFirst
                        
                        Do While Not frmAdVouchers.rsVouchers.EOF
                            mTotal = mTotal + frmAdVouchers.rsVouchers!amount
                            frmAdVouchers.rsVouchers.MoveNext
                        Loop
                        
                        frmAdVouchers.rsVouchers.MoveFirst
                        frmAdVouchers.txtTotal.Text = Format(mTotal, "#,##0.00")
                        
                    Else
                        
                        frmAdVouchers.txtNoOfRecords.Text = Format(frmAdVouchers.rsVouchers.RecordCount, "#,##0")
                        Lock_Button "TFFFTT", frmAdVouchers.cmdMenu, 5
                        frmAdVouchers.txtTotal.Text = "0.00"
                        
                    End If
                    
                End If
                
            Else
                MsgBox "No record to import.", vbExclamation + vbOKOnly
            End If
        End With
    Else
        MsgBox "No record found.", vbExclamation + vbOKOnly
    End If
    
End Sub

Private Sub fg_DblClick()
    With frmBrowseEmployee
        If fg.Col = 0 Then
            .mBrowseType = "Otherded"
            .Show vbModal
        End If
    End With
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If fg.Row >= 0 Then
            fg.RemoveItem fg.Row
        End If
    Else
    
    End If
End Sub

Private Sub Form_Load()

    dlg.DefaultExt = "xls"
    dlg.Filter = "Excel 97 (*.xls)|*.xls"
    
    fg.AllowUserResizing = flexResizeBoth
    fg.MergeCells = flexMergeSpill
    fg.ExtendLastCol = True
    
    If mImportType = "Otherded" Then
        lblOthers.Caption = "Other deductions : " & frmAdOtherDeductions.tdbOtherDeductions.Text
    ElseIf mImportType = "OtherEarnings" Then
        lblOthers.Caption = "Other earnings : " & frmAdOtherEarnings.tdbOtherEarnings.Text
    ElseIf mImportType = "Vouchers" Then
        lblOthers.Caption = ""
    End If
    
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
    
    Caption = "Import other deductions" & " [" & dlg.FileName & "]"

    sheet = 0
    btnNextSheet.Enabled = True
    btnPreviousSheet.Enabled = True
    
    fg.ColWidth(0) = 4500
    fg.ColWidth(1) = 500

End Sub

Private Sub btnNextSheet_Click()

    sheet = sheet + 1
    
    MousePointer = MousePointerConstants.vbHourglass
    
    On Error Resume Next
    fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    If err <> 0 Then
        sheet = sheet - 1
        fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    End If
    On Error GoTo 0
    
    MousePointer = MousePointerConstants.vbDefault

    'fg.ColWidth(0) = 4400
    'fg.ColWidth(1) = 500
    
End Sub

Private Sub btnPreviousSheet_Click()

    sheet = sheet - 1
    
    MousePointer = MousePointerConstants.vbHourglass
    
    On Error Resume Next
    fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    If err <> 0 Then
        sheet = sheet + 1
        fg.LoadGrid dlg.FileName, flexFileExcel, sheet
    End If
    On Error GoTo 0
    'fg.ColWidth(0) = 4500
    'fg.ColWidth(1) = 500
    MousePointer = MousePointerConstants.vbDefault

    
End Sub

Private Sub btnSave_Click()

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

