Attribute VB_Name = "modPointer"
Public Function pointme(flex As Object, txt2find As String, Col As Long)
On Error Resume Next
With flex
  If .Rows > 1 Then
        .Select .FindRow(txt2find, 1, Col, False, False), Col
        .ShowCell .FindRow(txt2find, 1, Col, False, False), Col
  End If
End With
End Function

Public Function updaterow(flex As Object, mRow As Long)
If mRow = flex.Rows Then
    With flex
        If flex.Row <> 1 Then
            .Row = mRow - 1
            If mRow <> 1 Then
                .ShowCell mRow - 1, 1
            End If
        End If
    End With
End If
If mRow <> 1 Then
    With flex
        If mRow <> 0 Then
            .Row = mRow - 1
            .ShowCell (mRow - 1), 1
        End If
    End With
End If
End Function

Public Function pointmetdg(grid As Object, rs As Recordset, fieldtosearch As String, txt2find As Integer)
With rs
    .Find fieldtosearch & " = " & txt2find & ""
End With
End Function
