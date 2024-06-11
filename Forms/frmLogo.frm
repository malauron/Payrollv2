VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin VB.Image imgLogo 
      Height          =   4620
      Left            =   -1020
      Picture         =   "frmLogo.frx":0000
      Top             =   900
      Width           =   8640
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    With imgLogo
        .Top = (Me.ScaleHeight / 2) - (.Height / 2)
        .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    End With
End Sub
