VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilInsertALILoan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utiliy"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   315
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      Caption         =   "&Insert"
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
      Image           =   "frmUtilInsertALILoan.frx":0000
      cBack           =   14737632
   End
End
Attribute VB_Name = "frmUtilInsertALILoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInsert_Click()

    Dim mLoanNo         As Integer
    Dim mLoanDedCode    As Integer
    
    Dim rsEmployee      As ADODB.Recordset
    
    NetOpen rsEmployee, "select * from employee where employeecode not in (374,147,189,220,229,235,216,16,325,183,8,33,35,42,45,46,48,55,56,58,66,70,91,68,78,99,102,108,129,133,162,163,165,171,179,185,186,194, " & _
                        "204,205,213,214,217,222,225,228,231,245,270,277,281,282,284,294,298,303,312,311,314,315,316,318,324,329,333,353,357,375,377,378,380)"
    
    With rsEmployee
        If .RecordCount > 0 Then
            
            .MoveFirst
            pb.Value = 0
            pb.Max = .RecordCount
            
            ConMain.Execute "set autocommit = 0 "
            ConMain.BeginTrans
            
            Do While Not .EOF
            
                pb.Value = pb.Value + 1
                mLoanNo = LastCode("Loans")
                
                ConMain.Execute "insert into loans(loancode,dummycode,employeecode,loantypescode,costcentercode," & _
                                    "divisioncode,branchcode,loandate,loanamnt,dedperpayday, " & _
                                    "noofinst,startdate,status,remarks,referenceno) values (" & _
                                    mLoanNo & ",'" & Format(mLoanNo, "0000000000") & "', " & !employeecode & ",9," & !costcentercode & ", " & _
                                    !divisioncode & "," & !branchcode & ",'2009-01-01',2538,105.75,12, " & _
                                    "'2009-01-01','Active','','')"
                
                mLoanDedCode = LastCode("LoanDed")
                
                ConMain.Execute "insert into loanded (loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled) values " & _
                              "(" & mLoanDedCode & "," & mLoanNo & ",9," & !employeecode & "," & _
                               0 & ",'" & Format(Now, "YYYY-MM-DD") & "', 0 ,2538,'Y','N')"
                
                mLoanDedCode = LastCode("LoanDed")
                
                ConMain.Execute "insert into loanded(loandedcode,loancode,loantypescode,employeecode,amtded,dateposted,ttlamtpaid,balance,fnlz,cancelled, " & _
                        "remarks) values " & _
                      "(" & mLoanDedCode & "," & mLoanNo & ",9," & !employeecode & ",1269,'" & Format(Now, "YYYY-MM-DD") & "', 1269,1269,'Y','N','')"
                       
                .MoveNext
            Loop
            
            ConMain.CommitTrans
            
            MsgBox "Import process completed!", vbInformation + vbOKOnly
            
        End If
    End With
End Sub
