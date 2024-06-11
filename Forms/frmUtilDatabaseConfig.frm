VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilDatabaseConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBText6Ctl.TDBText txtUserName 
      Height          =   285
      Left            =   5715
      TabIndex        =   1
      Top             =   1020
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":006C
      Key             =   "frmUtilDatabaseConfig.frx":008A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   330
      Left            =   5700
      TabIndex        =   6
      Top             =   2775
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   3186872
      cFHover         =   3186872
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   14737632
   End
   Begin TDBText6Ctl.TDBText txtPassword 
      Height          =   285
      Left            =   5715
      TabIndex        =   2
      Top             =   1365
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":00CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":013A
      Key             =   "frmUtilDatabaseConfig.frx":0158
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin lvButton.lvButtons_H cmdClose 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   7290
      TabIndex        =   7
      Top             =   2775
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      Caption         =   "&Close"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   3186872
      cFHover         =   3186872
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      cBack           =   14737632
   End
   Begin TDBText6Ctl.TDBText txtServer 
      Height          =   285
      Left            =   5715
      TabIndex        =   0
      Top             =   675
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":019C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":0208
      Key             =   "frmUtilDatabaseConfig.frx":0226
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtDatabase 
      Height          =   285
      Left            =   5715
      TabIndex        =   5
      Top             =   2400
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":026A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":02D6
      Key             =   "frmUtilDatabaseConfig.frx":02F4
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtConfirm 
      Height          =   285
      Left            =   5715
      TabIndex        =   3
      Top             =   1710
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":0338
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":03A4
      Key             =   "frmUtilDatabaseConfig.frx":03C2
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtPort 
      Height          =   285
      Left            =   5715
      TabIndex        =   4
      Top             =   2055
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   503
      Caption         =   "frmUtilDatabaseConfig.frx":0406
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilDatabaseConfig.frx":0472
      Key             =   "frmUtilDatabaseConfig.frx":0490
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Database Connection Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5445
      TabIndex        =   14
      Top             =   120
      Width           =   4080
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   13
      Top             =   1770
      Width           =   2730
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   12
      Top             =   2445
      Width           =   2730
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   11
      Top             =   2100
      Width           =   2730
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "MySQL Host Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   10
      Top             =   705
      Width           =   2730
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   9
      Top             =   1050
      Width           =   2730
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0030A0B8&
      Height          =   195
      Left            =   2895
      TabIndex        =   8
      Top             =   1425
      Width           =   2730
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3450
      Left            =   0
      Picture         =   "frmUtilDatabaseConfig.frx":04D4
      Top             =   0
      Width           =   9030
   End
End
Attribute VB_Name = "frmUtilDatabaseConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mString         As String

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdOK_Click()

    On Error GoTo err

    Dim Phrase      As String
    Dim FileName1   As String
    Dim txt         As String
    Dim Char1       As String
    
    Dim Position    As Integer
    
    Dim Asc1        As Long
    
    If Trim(txtServer.Text) = "" Then
        MsgBox "Please specify a server name.", vbExclamation + vbOKOnly
        txtServer.SetFocus
        Exit Sub
    End If
    
    If Trim(txtUserName.Text) = "" Then
        MsgBox "Please specify database username.", vbExclamation + vbOKOnly
        txtUserName.SetFocus
        Exit Sub
    End If

    If Trim(txtPassword.Text) = "" Then
        MsgBox "Please specify a password.", vbCritical
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text <> txtConfirm.Text Then
        MsgBox "Password confirmation do not match.", vbCritical
        txtPassword.SetFocus
        txtConfirm.Text = ""
        Exit Sub
    End If
    
    If Trim(txtPort.Text) = "" Then
        MsgBox "Please enter the MySQL port number.", vbExclamation + vbOKOnly
        txtPort.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDatabase.Text) = "" Then
        MsgBox "Please specify the database name.", vbExclamation + vbOKOnly
        txtDatabase.SetFocus
        Exit Sub
    End If
    
'encrypt
    Phrase = txtUserName.Text & "," & txtPassword.Text & "," & txtDatabase.Text & "," & txtServer.Text & "," & txtPort.Text
    mString = ""
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        Asc1 = Asc(Char1)
        Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
        Char1 = Chr$(Asc1)
        mString = mString & Char1
    Next
    'save
    If MsgBox("Are you sure you want to overwrite " & vbCrLf & _
                "the existing Database Login Information?", vbYesNo + vbQuestion) = vbYes Then
        'save
        FileName1 = App.Path & "\Extdata\m4ky.dat"
        If FileName1 = "" Then
            Exit Sub
        End If
        
        Open FileName1 For Output As #1
            Print #1, mString
        Close #1
        MsgBox "Password successfully changed. Please re-login.", vbInformation
        frmLogin.Show
        Unload Me
        
        Else
        Exit Sub
        Unload Me
    End If

    Exit Sub

err:
    If err.Number <> 0 Then
        MsgBox "An error occured. Please refer to the " & vbCrLf & _
               "following information." & vbCrLf & vbCrLf & _
               "Error Code: " & err.Number & vbCrLf & _
               "Error Description: " & err.Description & vbCrLf & _
               "Error Source: " & err.Source, vbCritical
        Exit Sub
    End If



End Sub

Private Sub txtConfirm_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDatabase_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPassword_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPort_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtServer_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtUserName_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
