Attribute VB_Name = "modAPI"
Option Explicit

'===============================
'API Declarations and Constant
'===============================

'For tracking mouse cursor position
Public Declare Function GetCursorPos Lib "user32" _
            (lpPoint As POINTAPI) As Long
            
Public Type POINTAPI
        X As Long
        Y As Long
End Type




