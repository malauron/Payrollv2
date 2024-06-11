Attribute VB_Name = "modMenuAPI"
' Declarations and such needed for the example:
' (Copy them to the (declarations) section of a module.)
Public Declare Function GetMenu Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Public Const MIIM_STATE = &H1
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFS_CHECKED = &H8

Public Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal _
  hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As _
  MENUITEMINFO) As Long


Public Declare Function GetWindowLong Lib "user32" Alias _
       "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias _
       "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long

Public Sub tvwInitialize(ByRef Control As MSComctlLib.TreeView)

    Const TVS_CHECKBOXES = &H100
    Const GWL_STYLE = (-16)
    
    Dim CurStyle As Long
    Dim Result As Long
    
    CurStyle = GetWindowLong(Control.hwnd, GWL_STYLE)
    Result = SetWindowLong(Control.hwnd, GWL_STYLE, _
             CurStyle Or TVS_CHECKBOXES)
    
End Sub

Public Sub tvwCheckBoxes(ByRef Control As MSComctlLib.TreeView, ByRef Node As MSComctlLib.Node)

    '-- childs
    tvwCheckBoxesAux1 Control, Node
    
    '-- parent
    tvwCheckBoxesAux2 Control, Node

End Sub

Private Sub tvwCheckBoxesAux1(ByRef Control As MSComctlLib.TreeView, ByRef Node As MSComctlLib.Node)

    Dim v As Boolean, i As Long, l As Long, f As Long, c As Long
    
    If Node.Children = 0 Then Exit Sub
    
    v = Node.Checked
    f = Node.Child.Index
    l = Node.Child.LastSibling.Index
    For i = f To l
        Control.Nodes(i).Checked = v
        tvwCheckBoxesAux1 Control, Control.Nodes(i)
    Next
    
End Sub

Private Sub tvwCheckBoxesAux2(ByRef Control As MSComctlLib.TreeView, ByRef Node As MSComctlLib.Node)

    Dim n As MSComctlLib.Node
    Dim v As Boolean, i As Long, l As Long, f As Long, c As Long
    
    If Node.Parent Is Nothing Then Exit Sub
    
    v = Node.Checked
    Set n = Node.Parent
    f = n.Child.Index
    l = n.Child.LastSibling.Index
    c = 0
    For i = f To l
        If Control.Nodes(i).Checked Then
            c = c + 1
        End If
    Next
    n.Checked = CBool(c)
    tvwCheckBoxesAux2 Control, n

End Sub

