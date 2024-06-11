Attribute VB_Name = "modOpenedForms"
Option Explicit

Public mForm_Count   As Integer

Public Sub Add_MDIButton(Form_Name As String, Form_Caption As String)

    Dim I           As Integer
    
    Dim mLeft       As Long
    Dim mLastNo     As Long
    
    Dim OldFrmName  As String
    Dim OldFrmCap   As String
    
    Dim mFound      As Boolean
    
    mForm_Count = mForm_Count + 1
    
    With mdiIdeasoftPayroll
        
        mFound = False
        OldFrmName = ""
        OldFrmCap = ""
        
        For I = 0 To .cmd.UBound
            If mFound = False Then
                If .cmd(I).Tag = Form_Name Then
                    .cmd(I).Visible = True
                    OldFrmName = .cmd(I).Tag
                    OldFrmCap = .cmd(I).Caption
                    mLastNo = I
                    mFound = True
                End If
            Else
                If .cmd(I).Visible = True Then
                    .cmd(mLastNo).Tag = .cmd(I).Tag
                    .cmd(mLastNo).Caption = .cmd(I).Caption
                    .cmd(I).Tag = OldFrmName
                    .cmd(I).Caption = OldFrmCap
                    mLastNo = I
                End If
            End If
        Next
        
        
        If mFound = False Then
            If .cmd(0).Visible = False Then
                .cmd(0).Left = 0
                .cmd(0).Top = 0
                .cmd(0).Width = .ScaleWidth
                .cmd(0).Visible = True
                .cmd(0).Caption = Form_Caption
                .cmd(0).Tag = Form_Name
            Else
                Load .cmd(.cmd.UBound + 1)
                .cmd(.cmd.UBound).ZOrder
                .cmd(.cmd.UBound).Tag = Form_Name
                .cmd(.cmd.UBound).Caption = Form_Caption
                .cmd(.cmd.UBound).Visible = True
            End If
        End If
        
        
        I = 0
        mLeft = 0
        For I = 0 To .cmd.UBound
            If .cmd(I).Visible = True Then
                .cmd(I).Top = 0
                .cmd(I).Width = (.ScaleWidth / (mForm_Count))
                .cmd(I).Left = mLeft
                mLeft = mLeft + .cmd(I).Width
            End If
        Next
        
    End With
End Sub

Public Sub Focus_MDIButton(mForm As Form)
    
    Dim I As Integer
    
    mForm.WindowState = vbMaximized
    
    With mdiIdeasoftPayroll
    
        For I = 0 To .cmd.UBound
        
            
                
            If .cmd(I).Tag <> mForm.Name Then
                
                .cmd(I).BackColor = &HE0E0E0
                .cmd(I).GradientColor = &HE0E0E0
                .cmd(I).HoverBackColor = &H80FFFF
                
            Else
            
                .cmd(I).BackColor = vbWhite
                .cmd(I).GradientColor = vbWhite
                .cmd(I).HoverBackColor = vbWhite
            
                mForm.ZOrder
                
            End If
            
        Next
        
    End With
    
End Sub

Public Sub Remove_MDIButton(Form_Name As String)

    Dim I           As Integer
    Dim mLeft       As Long
    
    mForm_Count = mForm_Count - 1
    
    With mdiIdeasoftPayroll
        For I = 0 To .cmd.UBound
            If .cmd(I).Tag = Form_Name Then
                .cmd(I).Visible = False
                Exit For
            End If
        Next
        
        I = 0
        mLeft = 0
        For I = 0 To .cmd.UBound
            If .cmd(I).Visible = True Then
                .cmd(I).Top = 0
                .cmd(I).Width = (.ScaleWidth / (mForm_Count))
                .cmd(I).Left = mLeft
                mLeft = mLeft + .cmd(I).Width
            End If
        Next
        
    End With
End Sub

