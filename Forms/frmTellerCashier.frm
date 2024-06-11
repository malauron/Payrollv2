VERSION 5.00
Begin VB.Form frmTellerCashier 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTellerCashier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   405
      Left            =   3870
      TabIndex        =   9
      Top             =   2940
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   405
      Left            =   900
      TabIndex        =   8
      Top             =   2940
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Teller or Verifier"
      Height          =   2655
      Left            =   3120
      TabIndex        =   4
      Top             =   90
      Width           =   2955
      Begin CitronSoftwarePayroll.ucTextBox txttvkbg 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txttvplya 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txttvrbnk 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1020
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txttvkbgExt 
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   1830
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txttvplyaExt 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   2220
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin VB.Label Label12 
         Caption         =   "For Extensions w/o Regular Bank"
         Height          =   345
         Left            =   60
         TabIndex        =   25
         Top             =   1560
         Width           =   2865
      End
      Begin VB.Label Label10 
         Caption         =   "Kaabag"
         Height          =   345
         Left            =   90
         TabIndex        =   23
         Top             =   1890
         Width           =   1125
      End
      Begin VB.Label Label9 
         Caption         =   "Pamilya"
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   2880
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Label Label6 
         Caption         =   "Regular Bank"
         Height          =   345
         Left            =   90
         TabIndex        =   15
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Pamilya"
         Height          =   345
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Kaabag"
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cashier"
      Height          =   2655
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2985
      Begin CitronSoftwarePayroll.ucTextBox txtcshrkbg 
         Height          =   375
         Left            =   1470
         TabIndex        =   1
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtcshrplya 
         Height          =   375
         Left            =   1470
         TabIndex        =   2
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtcshrrbnk 
         Height          =   375
         Left            =   1470
         TabIndex        =   3
         Top             =   1020
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtcshrkbgExt 
         Height          =   375
         Left            =   1470
         TabIndex        =   16
         Top             =   1830
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin CitronSoftwarePayroll.ucTextBox txtcshrplyaExt 
         Height          =   375
         Left            =   1470
         TabIndex        =   17
         Top             =   2220
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
      End
      Begin VB.Label Label11 
         Caption         =   "For Extensions w/o Regular Bank"
         Height          =   345
         Left            =   60
         TabIndex        =   24
         Top             =   1560
         Width           =   2865
      End
      Begin VB.Label Label8 
         Caption         =   "Kaabag"
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   1860
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Pamilya"
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   2880
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "Regular Bank"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Pamilya"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Kaabag"
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   6120
      Y1              =   2820
      Y2              =   2820
   End
End
Attribute VB_Name = "frmTellerCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
