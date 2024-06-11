VERSION 5.00
Begin VB.Form frmPayroll2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View employee payroll"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayroll2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbsdays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtAbsamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtLoanded 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtPlamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtPldays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtslamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtSldays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtVlamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtVldays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   72
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtNetpay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtDeductamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtTaxamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtHdmfamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtPhilamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtSSS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtGrosspay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtEarningsamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtNiteamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtNitehrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtOtsplxamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtOtsplxhrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtOtsplamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txtOtsplhrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtOtlegxamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtOtlegxhrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtOtlegamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtOtleghrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtOtsunxamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtOtsunxhrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtOtsunamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtOtsunhrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtOtregamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtOtreghrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtBasicpay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtSplamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtSpldays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtLegamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtLegdays 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtUtamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtUthrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtlateamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtlatehrs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtRegamnt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtDayswrk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtPayrate 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7230
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtRatetype 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   7230
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtEmpname 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtPerdesc 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label35 
      Caption         =   "Absences"
      Height          =   255
      Left            =   240
      TabIndex        =   83
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "Loan deductions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   82
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label33 
      Caption         =   "Others"
      Height          =   255
      Left            =   5640
      TabIndex        =   79
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label32 
      Caption         =   "SL"
      Height          =   255
      Left            =   5640
      TabIndex        =   76
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "VL"
      Height          =   255
      Left            =   5640
      TabIndex        =   73
      Top             =   1680
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9720
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label30 
      Caption         =   "Net pay"
      Height          =   255
      Left            =   5640
      TabIndex        =   71
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "Other deductions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   70
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label28 
      Caption         =   "Tax withheld"
      Height          =   255
      Left            =   5640
      TabIndex        =   65
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label27 
      Caption         =   "HDMF"
      Height          =   255
      Left            =   5640
      TabIndex        =   63
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label26 
      Caption         =   "Philhealth"
      Height          =   255
      Left            =   5640
      TabIndex        =   61
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "SSS"
      Height          =   255
      Left            =   5640
      TabIndex        =   59
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "DEDUCTIONS"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5640
      TabIndex        =   58
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label Label23 
      Caption         =   "Gross pay"
      Height          =   255
      Left            =   5640
      TabIndex        =   56
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label22 
      Caption         =   "Other earnings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   54
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "Night premium"
      Height          =   255
      Left            =   270
      TabIndex        =   51
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label20 
      Caption         =   "Special holiday(Ex)"
      Height          =   255
      Left            =   270
      TabIndex        =   48
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label19 
      Caption         =   "Special holiday"
      Height          =   255
      Left            =   270
      TabIndex        =   45
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label18 
      Caption         =   "Legal holiday(Ex)"
      Height          =   255
      Left            =   270
      TabIndex        =   42
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Legal holiday"
      Height          =   255
      Left            =   270
      TabIndex        =   39
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label16 
      Caption         =   "Sun/DayOff (Ex)"
      Height          =   255
      Left            =   270
      TabIndex        =   36
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "Sun/DayOff"
      Height          =   255
      Left            =   270
      TabIndex        =   33
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Regular"
      Height          =   255
      Left            =   270
      TabIndex        =   30
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "OVERTIME:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   270
      TabIndex        =   29
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Net basic pay"
      Height          =   255
      Left            =   270
      TabIndex        =   27
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Special holiday"
      Height          =   255
      Left            =   270
      TabIndex        =   24
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Legal holiday"
      Height          =   255
      Left            =   270
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Undertime"
      Height          =   255
      Left            =   270
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Late"
      Height          =   255
      Left            =   270
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Regular pay"
      Height          =   255
      Left            =   270
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Days worked"
      Height          =   255
      Left            =   270
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "EARNINGS"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   270
      TabIndex        =   8
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Rate"
      Height          =   255
      Left            =   6150
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Rate type"
      Height          =   255
      Left            =   6150
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Employee name"
      Height          =   255
      Left            =   390
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Payroll period"
      Height          =   255
      Left            =   390
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmPayroll2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
