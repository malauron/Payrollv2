VERSION 5.00
Begin VB.Form frmUsers2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add user"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2265
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2265
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtFullname 
      Height          =   315
      Left            =   2265
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtUserid 
      Height          =   315
      Left            =   2265
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label mAdd 
      Caption         =   "Y"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Caption         =   "Confirm password"
      Height          =   255
      Left            =   345
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   345
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Full name"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmUsers2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
