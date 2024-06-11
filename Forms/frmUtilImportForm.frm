VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilImportAll 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   165
      Left            =   30
      TabIndex        =   1
      Top             =   435
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2220
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   165
      Left            =   30
      TabIndex        =   2
      Top             =   630
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb3 
      Height          =   165
      Left            =   30
      TabIndex        =   3
      Top             =   825
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb4 
      Height          =   165
      Left            =   30
      TabIndex        =   4
      Top             =   1020
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb5 
      Height          =   165
      Left            =   30
      TabIndex        =   5
      Top             =   1215
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb6 
      Height          =   165
      Left            =   30
      TabIndex        =   6
      Top             =   1410
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmUtilImportAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdImport_Click()

    Dim rsm201          As ADODB.Recordset
    
    Dim mDivisionCode   As Integer
    Dim mCostCenterCode As Integer
    Dim mJobTitleCode   As Integer
    
    Dim mEmpNo          As Integer
    
    Dim mCivilStatus    As String
    Dim mBirthDate      As String
    Dim mDateHired      As String
    
    CitronPayroll.Execute "set autocommit =0"
    CitronPayroll.BeginTrans
    CitronPayroll.Execute "delete from branch"
    CitronPayroll.Execute "delete from division"
    CitronPayroll.Execute "delete from costcenter"
    CitronPayroll.Execute "delete from jobtitle"
    CitronPayroll.Execute "delete from employee"
    CitronPayroll.Execute "delete from lastcodeseries where module in ('Branch','CostCenter','Division','JobTitle','Employee')"
    CitronPayroll.Execute "insert into Branch values (" & LastCode("GetLastCodeA", "Branch") & ",'00000001', 'Cebu Branch', '')"
    CitronPayroll.Execute "update m201 set mname = '' where mname is Null"
    
    NetOpen rsm201, "", "select div_code from m201 group by div_code order by div_code"
    
    If rsm201.RecordCount > 0 Then
        rsm201.MoveFirst
        pb1.Max = rsm201.RecordCount
        pb1.Min = 0
        Do While Not rsm201.EOF
            pb1.Value = pb1.Value + 1
            mDivisionCode = LastCode("GetLastCodeA", "Division")
            CitronPayroll.Execute "insert into Division values (1, " & mDivisionCode & ",'" & Format(mDivisionCode, "00000000") & "', '" & rsm201!div_code & "', '')"
            CitronPayroll.Execute "update m201 set divisioncode = " & mDivisionCode & " where div_code = '" & rsm201!div_code & "'"
            rsm201.MoveNext
        Loop
    End If
    
    NetOpen rsm201, "", "select x1.dept_code,(select divisioncode from m201 where divisioncode = x1.divisioncode limit 1) divisioncode from m201 x1 group by dept_code order by x1.dept_code "
    
    If rsm201.RecordCount > 0 Then
        rsm201.MoveFirst
        pb2.Max = rsm201.RecordCount
        pb2.Min = 0
        Do While Not rsm201.EOF
            pb2.Value = pb2.Value + 1
            mCostCenterCode = LastCode("GetLastCodeA", "CostCenter")
            CitronPayroll.Execute "insert into CostCenter values (1, " & rsm201!divisioncode & "," & mCostCenterCode & ",'" & Format(mCostCenterCode, "00000000") & "', '" & rsm201!dept_code & "', '',0)"
            CitronPayroll.Execute "update m201 set costcentercode = " & mCostCenterCode & " where dept_code = '" & rsm201!dept_code & "'"
            rsm201.MoveNext
        Loop
    End If
    
    NetOpen rsm201, "", "select position from m201 group by position order by position"
    
    If rsm201.RecordCount > 0 Then
        rsm201.MoveFirst
        pb3.Max = rsm201.RecordCount
        pb3.Min = 0
        Do While Not rsm201.EOF
            pb3.Value = pb3.Value + 1
            mJobTitleCode = LastCode("GetLastCodeA", "JobTitle")
            CitronPayroll.Execute "insert into JobTitle values (" & mJobTitleCode & ", '" & Format(mJobTitleCode, "00000000") & "', '" & rsm201!Position & "', '')"
            CitronPayroll.Execute "update m201 set jobtitlecode = " & mJobTitleCode & " where position = '" & rsm201!Position & "'"
            rsm201.MoveNext
        Loop
    End If
    
    NetOpen rsm201, "", "select * from m201 order by lname,fname,mname"

    If rsm201.RecordCount > 0 Then
        rsm201.MoveFirst

        pb4.Max = rsm201.RecordCount
        pb4.Min = 0
        
        Do While Not rsm201.EOF
        
            pb4.Value = pb4.Value + 1

            If rsm201!civil = "S" Then
                mCivilStatus = "Single"
            ElseIf rsm201!civil = "M" Then
                mCivilStatus = "Married"
            ElseIf rsm201!civil = "W" Then
                mCivilStatus = "Widow"
            ElseIf rsm201!civil = "D" Then
                mCivilStatus = "Divorced"
            Else
                mCivilStatus = "Single"
            End If
            
            If Not IsDate(rsm201!birthday) Then
                mBirthDate = Format(Now, "YYYY-MM-DD")
            Else
                mBirthDate = Format(rsm201!birthday, "YYYY-MM-DD")
            End If
            
            If Not IsDate(rsm201!date_emp) Then
                mDateHired = Format(Now, "YYYY-MM-DD")
            Else
                mDateHired = Format(rsm201!date_emp, "YYYY-MM-DD")
            End If
            
            mEmpNo = LastCode("GetLastCodeA", "Employee")

            CitronPayroll.Execute "insert into employee (empno,dummycode,biometid,lastname,firstname,middlename,gender,civilstatus,birthdate, " & _
                          "houseno,street,provcode,muncode,brgycode,branchcode,divisioncode,costcentercode,telno,mobileno,email, " & _
                          "emrgncyname,emrgncyno,emrgncyemail,payfreqcode, " & _
                          "wtcode,jobtitlecode,empstatcode,ratetypecode,monthly_rate,daily_rate,hourly_rate, " & _
                          "sssno,philhno,tinno,hdmfno,bankacctno, " & _
                          "sectioncode,sssamt,ssser,sssec,philhamt,philher,taxamt,hdmfamt,hdmfer, " & _
                          "sssauto,philhauto,taxauto,hdmfauto,saltobank," & _
                          "regular,isactive,logbased,mealallow,fixedEarnings,datehired) values " & _
                          "(" & mEmpNo & ",'" & Format(mEmpNo, "00000000") & "','','" & UCase(rsm201!lname) & "','" & UCase(rsm201!fname) & "','" & UCase(rsm201!mname) & "','" & IIf(rsm201!sex = "F", "Female", "Male") & "','" & mCivilStatus & "','" & mBirthDate & "', " & _
                          "'','',Null,Null,Null,1," & rsm201!divisioncode & "," & rsm201!costcentercode & ",'','','', " & _
                          "'','','',3," & _
                          "''," & rsm201!jobtitlecode & "," & IIf(rsm201!emp_status = "S", 1, 2) & "," & IIf(rsm201!emp_status = "S", 4, 1) & ", " & rsm201!month_pay & ", " & rsm201!daily & ", " & rsm201!hourly & ", " & _
                          "" & IIf(IsNull(rsm201!SSS), "Null", "'" & rsm201!SSS & "'") & "," & IIf(IsNull(rsm201!philno), "Null", "'" & rsm201!philno & "'") & "," & IIf(IsNull(rsm201!tin), "Null", "'" & rsm201!tin & "'") & "," & IIf(IsNull(rsm201!Pagibig), "Null", "'" & rsm201!Pagibig & "'") & ",'', " & _
                          "'',0,0,0,0,0,0,0,0, " & _
                          "1,1,1,1, 'Y'," & _
                          "'" & IIf(rsm201!emp_status = "S", "Y", "N") & "', 'Y', 'Y',0,0,'" & mDateHired & "')"
            CitronPayroll.Execute "update m201 set empno = " & mEmpNo & " where lname = '" & rsm201!lname & "' and fname = '" & rsm201!fname & "' and mname = '" & rsm201!mname & "'"

            rsm201.MoveNext

        Loop

    End If
    
    CitronPayroll.CommitTrans
    

End Sub
