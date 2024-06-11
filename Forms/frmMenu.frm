VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   9480
   Tag             =   "Shortcut Menu"
   WindowState     =   2  'Maximized
   Begin VB.Frame fra1 
      BackColor       =   &H00E0E0E0&
      Height          =   7995
      Left            =   -15
      TabIndex        =   0
      Top             =   -105
      Width           =   9510
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Add_MDIButton Me.Name, Me.Tag

End Sub
Private Sub Form_Resize()
    
    Me.WindowState = 2
    
    With fra1
        .Top = (Me.ScaleHeight / 2) - (.Height / 2)
        .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    End With

End Sub
