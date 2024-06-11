VERSION 5.00
Begin VB.Form frmScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.Image img1 
      Height          =   1080
      Left            =   7485
      Picture         =   "frmScreen.frx":0000
      Top             =   4815
      Width           =   7560
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    With img1
        .Top = Me.ScaleHeight - .Height
        .Left = Me.ScaleWidth - .Width
    End With
End Sub
