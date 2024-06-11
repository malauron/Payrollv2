VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMDUsergroups_Restrictions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Restrictions"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      Height          =   6510
      Left            =   30
      TabIndex        =   3
      Top             =   -75
      Width           =   7815
      Begin MSComctlLib.TreeView tvwGroups 
         Height          =   6300
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   11113
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   7815
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   60
         TabIndex        =   1
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmMDUsergroups_Restrictions.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   1800
         TabIndex        =   0
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   688
         Caption         =   "&OK"
         CapAlign        =   2
         BackStyle       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "frmMDUsergroups_Restrictions.frx":0CDA
         cBack           =   14737632
      End
   End
End
Attribute VB_Name = "frmMDUsergroups_Restrictions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public userGroupID As Integer

' This code is licensed according to the terms and conditions listed here.


' When button Command1 is pressed, output the structure of the entire menu system
' of Form1 to the Debug window.  The entire menu heirarchy is displayed, and any items
' that are checked are identified.  A recursive subroutine is used to output the contents
' of each individual menu, calling itself whenever a submenu is found.

' *** Place the following code inside a module. ***
' This function performs the recursive output of the menu structure.
Private Sub IterateThroughItems(ByVal hMenu As Long, ByVal level As Long, lastID As String)

  Dim n As Node
  Dim menuCaption As String

  ' hMenu is a handle to the menu to output
  ' level is the level of recursion, used to indent submenu items
  Dim itemcount As Long    ' the number of items in the specified menu
  Dim c As Long            ' loop counter variable
  Dim mii As MENUITEMINFO  ' receives information about each item
  Dim retval As Long       ' return value
  
  ' Count the number of items in the menu passed to this subroutine.
  itemcount = GetMenuItemCount(hMenu)
  
  ' Loop through the items, getting information about each one.
  With mii
    .cbSize = Len(mii)
    .fMask = MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU
    For c = 0 To itemcount - 1
      ' Make room in the string buffer.
      .dwTypeData = Space(256)
      .cch = 256
      ' Get information about the item.
      retval = GetMenuItemInfo(hMenu, c, 1, mii)
      ' Output a line of information about this item.
      
      If mii.fType = MFT_SEPARATOR Then
        ' This is a separator bar.
        'Debug.Print "   " & String(3 * level, ".") & "-----"
      Else
        ' This is a text item.
        ' If this is checked, display (X) in the margin.
        'Debug.Print IIf(.fState And MFS_CHECKED, "(X)", "   ");
        ' Display the text of the item.
        'Debug.Print String(3 * level, ".") & Left(.dwTypeData, .cch)
        ' If this item opens a submenu, display its contents.
        menuCaption = Swap(Left(.dwTypeData, .cch))
        If .hSubMenu <> 0 Then
          
          If lastID = "a" Then
            Set n = tvwGroups.Nodes.Add(Key:=lastID & level & c, Text:=menuCaption)
          Else
            Set n = tvwGroups.Nodes.Add(lastID, tvwChild, lastID & level & c, menuCaption)
          End If
          IterateThroughItems .hSubMenu, level + 1, lastID & level & c
        Else
          Set n = tvwGroups.Nodes.Add(lastID, tvwChild, lastID & level & c, menuCaption)
        End If
        
        
      End If
    Next c
  End With
End Sub

' *** Place the following code inside Form1. ***
' When Command1 is clicked, output the entire contents of Form1's menu system.
Private Sub Command1_Click()
  Dim hMenu As Long  ' handle to the menu bar of Form1
  
  ' Get a handle to Form1's menu bar.
  hMenu = GetMenu(mdiIdeasoftPayroll.hwnd)
  ' Use the above function to output its contents.
  IterateThroughItems hMenu, 0, "a"
End Sub

Private Sub Form_Load()
    Dim hMenu As Long  ' handle to the menu bar of Form1
  
    ' Get a handle to Form1's menu bar.
    hMenu = GetMenu(mdiIdeasoftPayroll.hwnd)
    ' Use the above function to output its contents.
    IterateThroughItems hMenu, 0, "a"
    
    Dim rsRestrictions As New ADODB.Recordset
    
    NetOpen rsRestrictions, "select * from usergroup_restrictions where usergroup_id = " & userGroupID
  
    Dim i As Integer
    With rsRestrictions
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                i = 1
                For i = 1 To tvwGroups.Nodes.count
                  If tvwGroups.Nodes(i).Text = !menuCaption Then
                    tvwGroups.Nodes(i).Checked = True
                    Exit For
                  End If
                Next
                .MoveNext
            Loop
        End If
    End With
End Sub

'
'Private Sub Command2_Click()
''    IterateChildren tvwGroups.Nodes(1)
'Dim i As Integer
'For i = 1 To tvwGroups.Nodes.count
'  If tvwGroups.Nodes(i).Checked = True Then
'    Dim c As Control
'    For Each c In Controls
'        If TypeOf c Is Menu Then
'            If c.Caption = tvwGroups.Nodes(i).Text Then
'              MsgBox c.Name
'            End If
'        End If
'    Next
'  End If
'Next
'
'End Sub

Private Sub tvwGroups_NodeCheck(ByVal Node As MSComctlLib.Node)
    tvwCheckBoxes tvwGroups, Node
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrHndlr
    ConMain.Execute "set autocommit = 0"
    ConMain.BeginTrans
    ConMain.Execute "delete from usergroup_restrictions " & _
                    "where usergroup_id = " & userGroupID & " and module = '" & ModuleVersion & "'"
    
    Dim c As Control
    Dim i As Integer
    For Each c In mdiIdeasoftPayroll.Controls
      If TypeOf c Is Menu Then
        i = 1
        For i = 1 To tvwGroups.Nodes.count
          If tvwGroups.Nodes(i).Text = Swap(c.Caption) Then
            If tvwGroups.Nodes(i).Checked Then
              ConMain.Execute "insert into usergroup_restrictions (usergroup_id,menuname,menucaption,module) values (" & _
                              userGroupID & ",'" & c.Name & "','" & tvwGroups.Nodes(i).Text & "','old_payroll')"
              Exit For
            End If
          End If
        Next
      End If
    Next
    ConMain.CommitTrans
    Unload Me
        
    Exit Sub
ErrHndlr:
    
    MsgBox "Error Message: " & err.Description, vbCritical + vbOKOnly
    
End Sub

Private Sub Form_Activate()
'  txtUserPassword.SetFocus
End Sub

Private Sub txtUserPassword_GotFocus()
'  With txtUserPassword
'    .SelStart = 0
'    .SelLength = Len(.Text)
'  End With
End Sub

Private Sub txtConfirmPassword_GotFocus()
'    With txtConfirmPassword
'      .SelStart = 0
'      .SelLength = Len(.Text)
'    End With
End Sub
  
