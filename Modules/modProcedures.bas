Attribute VB_Name = "modProcedures"
Option Explicit

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
        Left        As Long
        Top         As Long
        Right       As Long
        Bottom      As Long
End Type

Public Function Swap(mString As String) As String
    mString = Replace(mString, "'", "''")
    mString = Replace(mString, "\", "'\'")
    mString = Replace(mString, "&", "")
    Swap = mString
End Function

Sub NetOpen(ByRef rs As ADODB.Recordset, ByRef mSql As String, Optional mCon As ADODB.Connection, Optional mCursorType As CursorTypeEnum, Optional mLockType As LockTypeEnum)
    
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    If mCon Is Nothing Then
        Set mCon = ConMain
    End If
        
    If mCursorType <= 0 Then
        mCursorType = adOpenKeyset
    End If
    
    If mLockType <= 0 Then
        mLockType = adLockOptimistic
    End If

    rs.Open mSql, mCon, mCursorType, mLockType

End Sub

Public Sub bind_tdb(mCon As ADODB.Connection, ByRef srcDC As TDBCombo, ByVal srcSQL As String, ByVal srcFld As String, ByVal srcColBound As String, Optional ShowFirstRec As Boolean)
    
    Dim rs As New Recordset
    Dim Col As TrueOleDBList80.Column
    rs.CursorLocation = adUseClient
    
    rs.Open srcSQL, mCon, adOpenStatic, adLockOptimistic
    With srcDC
      .BoundText = ""
      .RowSource = rs
      .ListField = srcFld
      .BoundColumn = srcColBound
      .Columns(0).DataField = srcColBound
      .Columns(1).DataField = srcFld
      If rs.RecordCount > 0 And ShowFirstRec Then
        .BoundText = rs.Fields(srcColBound)
      End If
    End With
    
    Set rs = Nothing
    
End Sub

Public Function SearchList(KeyAscii As Integer, obj1 As Object, rs As ADODB.Recordset, sStr As String)
'Purpose : This routine allows the user to Automatically search from the List(Data Combo/Data List) when keypress is made in Data Combo Box
'On Error Resume Next
Dim iLen, iStart, iSelLength As Integer
Dim sCriteria As String
Dim tempstr As String

    'If control was locked then exit
'    If obj1.Locked Then
'      Exit Function
'    End If
    
    If KeyAscii = 8 Then
        tempstr = obj1.Text
        iStart = obj1.SelStart
        iSelLength = obj1.SelLength
        obj1.Text = tempstr
        If iStart > 0 Then
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength + 1
           Else
            obj1.SelStart = iStart
            obj1.SelLength = iSelLength
        End If
        
        KeyAscii = 0
        Exit Function
    End If

    If KeyAscii = 27 Then KeyAscii = 0: Exit Function
    If KeyAscii = 22 Then KeyAscii = 0:  Exit Function
    If Chr(KeyAscii) = "'" Then KeyAscii = Asc("`")
    If Not printable(KeyAscii) Then Exit Function

    iStart = obj1.SelStart + 1
    obj1.SelText = Chr(KeyAscii)
    KeyAscii = 0
    
    tempstr = sStr
    
    If rs.RecordCount = 0 Then
            obj1.Text = tempstr
            iStart = obj1.SelStart + 1
            iSelLength = obj1.SelLength
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength
            KeyAscii = 0
            Exit Function
    End If
    
    iLen = obj1.SelStart + 1
    iSelLength = obj1.SelLength
    
    If KeyAscii <> 1 Then obj1.SelText = Chr(KeyAscii)
      With rs
              .MoveFirst
              sCriteria = obj1.ListField & " like '" & obj1.Text & "%'"
              .Find sCriteria
              If Not .EOF Then
                 obj1.Text = Trim(.Fields(obj1.ListField) & " ")
                 obj1.SelStart = iStart
                 iLen = Len(obj1.Text)
                 If iLen = 0 Then Exit Function
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              Else
                  If tempstr <> "" Then
                      .MoveFirst
                      sCriteria = obj1.ListField & " like '" & obj1.Text & "%'"
                      .Find sCriteria
                  End If
                  If .AbsolutePosition > 0 Then
                      obj1.Text = Trim(.Fields(obj1.ListField) & " ")
                  Else
                      obj1.Text = tempstr
                  End If
                 obj1.SelStart = iStart - 1
                 iLen = Len(obj1.Text)
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              End If
      End With
End Function

Public Function SearchRecord(KeyAscii As Integer, obj1 As Object, rs As ADODB.Recordset, sStr As String, mSortBy As String)
    
    'Purpose : This routine searches a record typed in the textbox into the tdbgrid.
    On Error Resume Next
    Dim iLen, iStart, iSelLength As Integer
    Dim sCriteria As String
    Dim tempstr As String

    'If control was locked then exit
'    If obj1.Locked Then
'      Exit Function
'    End If
    
    If KeyAscii = 8 Then
        tempstr = obj1.Text
        iStart = obj1.SelStart
        iSelLength = obj1.SelLength
        obj1.Text = tempstr
        If iStart > 0 Then
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength + 1
           Else
            obj1.SelStart = iStart
            obj1.SelLength = iSelLength
        End If
        
        KeyAscii = 0
        Exit Function
    End If

    If KeyAscii = 27 Then KeyAscii = 0: Exit Function
    If KeyAscii = 22 Then KeyAscii = 0:  Exit Function
    If Chr(KeyAscii) = "'" Then KeyAscii = Asc("`")
    If Not printable(KeyAscii) Then Exit Function

    iStart = obj1.SelStart + 1
    obj1.SelText = Chr(KeyAscii)
    KeyAscii = 0
    
    tempstr = sStr
    
    If rs.RecordCount = 0 Then
            obj1.Text = tempstr
            iStart = obj1.SelStart + 1
            iSelLength = obj1.SelLength
            obj1.SelStart = iStart - 1
            obj1.SelLength = iSelLength
            KeyAscii = 0
            Exit Function
    End If
    
    iLen = obj1.SelStart + 1
    iSelLength = obj1.SelLength
    
    If KeyAscii <> 1 Then obj1.SelText = Chr(KeyAscii)
      With rs
              .MoveFirst
              sCriteria = mSortBy & " like '" & obj1.Text & "%'"
              .Find sCriteria
              If Not .EOF Then
                 obj1.Text = Trim(.Fields(mSortBy) & " ")
                 obj1.SelStart = iStart
                 iLen = Len(obj1.Text)
                 If iLen = 0 Then Exit Function
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              Else
                  If tempstr <> "" Then
                      .MoveFirst
                      sCriteria = mSortBy & " like '" & tempstr & "%'"
                      .Find sCriteria
                      If .EOF Then
                        .MoveFirst
                      End If
                  Else
                    .MoveFirst
                  End If
                  If .AbsolutePosition > 0 Then
                      obj1.Text = Trim(.Fields(mSortBy) & " ")
                  Else
                      obj1.Text = tempstr
                  End If
                 obj1.SelStart = iStart - 1
                 iLen = Len(obj1.Text)
                 obj1.SelLength = Len(obj1.Text) - iStart + 1
              End If
      End With
      
End Function

Private Function printable(ch As Integer) As Boolean
  
  Dim chrs As String
  
  chrs = " ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890`~!@#$%^&*()_+-={}[]:;""',./\<>?|"
  If InStr(chrs, UCase(Chr(ch))) > 0 Then printable = True Else printable = False
  
End Function

'Locks buttons
Public Sub Lock_Button(TF As String, Obj As Object, MaxNum As Integer)

  Dim i As Integer
  
  For i = 0 To MaxNum
    Obj(i).Enabled = IIf(Mid(TF, i + 1, 1) = "T", True, False)
  Next
  
End Sub

'Locks Tabs
Public Sub Lock_Tab(TF As String, Obj As Object, MaxNum As Integer)
  
  Dim i As Integer
  
  For i = 0 To MaxNum
    Obj.TabEnabled(i) = IIf(Mid(TF, i + 1, 1) = "T", True, False)
  Next
  
End Sub

'Locks Frames
Public Sub Lock_Frame(TF As String, Obj As Object, MaxNum As Integer)

  Dim i As Integer
  
  For i = 0 To MaxNum
    Obj(i).Enabled = IIf(Mid(TF, i + 1, 1) = "T", True, False)
  Next

End Sub

Public Sub CreateTmpDB(ByRef rs As ADODB.Recordset)
  Set rs = New ADODB.Recordset
  With rs
    .Fields.Append "Code", adVarChar, 50
    .Fields.Append "Description", adVarChar, 100
    .Open
  End With
End Sub

Public Sub Bind_tdd(mCon As ADODB.Connection, tddObj As Object, mSql As String, mListfield As String)
  
  Dim rsFn As ADODB.Recordset
  Set rsFn = New ADODB.Recordset
  
  rsFn.CursorLocation = adUseClient
  rsFn.Open mSql, mCon, adOpenStatic, adLockOptimistic
  
  With tddObj
    .DataSource = rsFn
    .ListField = mListfield
  End With
  
End Sub

Public Function DirExists(DirName As String) As Boolean
    
    On Error GoTo errorHandler
    
    DirExists = False
    'test the directory attribute
    If Dir(DirName) <> "" Then
        DirExists = True
    End If
    
    Exit Function
    
errorHandler:
    ' if an error occurs, this function returns False
End Function


Public Sub DestroyAllObjects()

Dim lo_Form As VB.Form

On Error GoTo Routine_Error

For Each lo_Form In VB.Forms
    If lo_Form.Name = "mdiIdeasoftPayroll" Then
    'skip unload
    Else
        Unload lo_Form
    End If
Next lo_Form

Routine_Exit:
Set lo_Form = Nothing

Exit Sub

Routine_Error:
'ms_ErrLocation = ms_MODULE & "DestroyAllObjects"
'AppLog.HandleError err.Number, err.Description, ms_ErrLocation
GoTo Routine_Exit

End Sub

