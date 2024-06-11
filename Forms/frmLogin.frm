VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3450
   ClientLeft      =   5910
   ClientTop       =   4890
   ClientWidth     =   9030
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2038.373
   ScaleMode       =   0  'User
   ScaleWidth      =   8478.681
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2670
      Left            =   1920
      TabIndex        =   0
      Top             =   90
      Width           =   7035
      Begin TDBText6Ctl.TDBText txtUsername 
         Height          =   285
         Left            =   2490
         TabIndex        =   1
         Top             =   1290
         Width           =   4035
         _Version        =   65536
         _ExtentX        =   7117
         _ExtentY        =   503
         Caption         =   "frmLogin.frx":6852
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLogin.frx":68BE
         Key             =   "frmLogin.frx":68DC
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
      Begin TDBText6Ctl.TDBText txtPassword 
         Height          =   285
         Left            =   2490
         TabIndex        =   2
         Top             =   1680
         Width           =   4035
         _Version        =   65536
         _ExtentX        =   7117
         _ExtentY        =   503
         Caption         =   "frmLogin.frx":6920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLogin.frx":698C
         Key             =   "frmLogin.frx":69AA
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
      Begin VB.Label lblDatabase 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Database Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2400
         Width           =   3270
      End
      Begin VB.Label lblServer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   2130
         Width           =   3270
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1005
         TabIndex        =   5
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   990
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLogin.frx":69EE
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Width           =   6810
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   600
         Picture         =   "frmLogin.frx":6AA1
         Top             =   1305
         Width           =   720
      End
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   330
      Left            =   5925
      TabIndex        =   6
      Top             =   2820
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmLogin.frx":4F32B
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H cmdClose 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   7470
      TabIndex        =   7
      Top             =   2820
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      Caption         =   "E&XIT"
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
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "frmLogin.frx":4FC05
      cBack           =   14737632
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   75
      Picture         =   "frmLogin.frx":504DF
      Top             =   -30
      Width           =   1785
   End
   Begin VB.Label lblCompanyInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1905
      TabIndex        =   8
      Top             =   3180
      Width           =   7020
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mEncrypted      As String
Dim mUsername       As String
Dim mPassword       As String
Dim mServerName     As String
Dim mDatabase       As String
Dim mPort           As String

Dim CTR             As Integer

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdOK_Click()

    Dim rsUser      As ADODB.Recordset
    
    If Trim(txtUsername.Text) = "" Then
        MsgBox "Please enter your username.", vbExclamation + vbOKOnly
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Please enter your password.", vbExclamation + vbOKOnly
        txtPassword.SetFocus
        Exit Sub
    End If
    
    NetOpen rsUser, "select * from users where username = '" & Swap(txtUsername.Text) & "'"
    
    If rsUser.RecordCount > 0 Then
        
        GlobalUserID = rsUser!user_id
        UserName = rsUser!UserName
        GlobalUserGroupID = rsUser!usergroup_id
        
        NetOpen rsUser, "select * from users where username = '" & Swap(UserName) & "' and password = PASSWORD('" & Swap(txtPassword.Text) & "')"
        If rsUser.RecordCount > 0 Then
            'UserType = rsUser!UserType
            mdiIdeasoftPayroll.Show
            mLogOn = True
            Unload Me
        Else
            MsgBox "Invlid password.", vbExclamation + vbOKOnly
            txtPassword.SetFocus
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
            Exit Sub
        End If
    Else
        MsgBox "Username not found.", vbExclamation + vbOKOnly
        txtUsername.SetFocus
        txtUsername.SelStart = 0
        txtUsername.SelLength = Len(txtUsername.Text)
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()

    On Error GoTo errH
    
    Dim rsUsers         As ADODB.Recordset
    Dim rsCompanyInfo   As ADODB.Recordset
    Dim DriverODBC      As String
    
    Dim FileName1       As String
    Dim txt             As String
    Dim Phrase          As String
    Dim Char1           As String
    
    Dim Position        As Integer
    
    Dim Asc1            As Long
    
    ModuleVersion = "old_payroll"
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    If Dir(App.Path & "\ExtData", vbDirectory) = vbNullString Then
            MkDir (App.Path & "\ExtData")
            If Dir(App.Path & "\ExtData\m4ky.dat", vbNormal) = vbNullString Then
                Open App.Path & "\ExtData\m4ky.dat" For Output As #1
                    Print #1, ""
                Close #1
            Else
                FileName1 = App.Path & "\ExtData\m4ky.dat"
            End If
    Else
        'false
        If Dir(App.Path & "\ExtData\m4ky.dat", vbNormal) = vbNullString Then
            Open App.Path & "\ExtData\m4ky.dat" For Output As #1
                Print #1, ""
            Close #1
            FileName1 = App.Path & "\ExtData\m4ky.dat"
        End If
    End If
    
    FileName1 = App.Path & "\ExtData\m4ky.dat"
    
    mEncrypted = ""
    
    Open FileName1 For Input As #1
        
        Do Until EOF(1)
            Input #1, txt
            mEncrypted = mEncrypted & txt
        Loop
        
    Close #1
    
    Phrase = mEncrypted
    
    mEncrypted = ""
    
    CTR = 0
    
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        Asc1 = Asc(Char1)
        Asc1 = (((Asc1 * Asc1) / 2) / 2)
        Asc1 = Sqr(Asc1)
        Char1 = Chr$(Asc1)
        '----
        If Char1 = "," Then
            Char1 = ""
            CTR = CTR + 1
            If CTR = 1 Then
                mUsername = mEncrypted
                mEncrypted = ""
            ElseIf CTR = 2 Then
                mPassword = mEncrypted
                mEncrypted = ""
            ElseIf CTR = 3 Then
                mDatabase = mEncrypted
                mEncrypted = ""
            ElseIf CTR = 4 Then
                mServerName = mEncrypted
                mEncrypted = ""
            End If
        End If
        '-----
        mEncrypted = mEncrypted & Char1
    Next
    
    mPort = mEncrypted
    
    mEncrypted = mEncrypted
    
    SQLServerName = mServerName
    SQLDatabase = mDatabase
    SQLUsername = mUsername
    SQLPassword = mPassword
    SQLPort = mPort
    
    If ConMain.State <> 0 Then ConMain.Close 'Check if currently connected if yes, disconnect.
    
    ConMain.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & SQLServerName & ";DATABASE=" & SQLDatabase & ";" _
                                 & "UID=" & SQLUsername & ";PWD=" & SQLPassword & "; PORT=" & SQLPort & "; OPTION=3"
    
    ConMain.Open  'open the connection
    
'    If (MySQLDSNWanted(SQLDatabase)) = False Then
'        DriverODBC = String(255, Chr(32))
'        If Not MakeMySQLDSN(DriverODBC, SQLDatabase) Then
'            MsgBox "Error Occured-DSN  could  not  be  Created. " & vbCrLf & _
'                   "Please notify software vendor immediately.", vbOKOnly + vbCritical, "Error...!!"
'        End If
'    End If
    
    mEmpPicPath = App.Path & "\Images"
    
    lblServer.Caption = "Server Name : " & mServerName
    lblDatabase.Caption = "Database Name : " & mDatabase
    
    NetOpen rsUsers, "select * from users limit 1"
    If rsUsers.RecordCount = 0 Then
        ConMain.Execute "insert into users(fullname,username,password) values ('LinkPro Technologies Inc.','linkpro','linkpro')"
    End If
    
     NetOpen rsCompanyInfo, "select * from companyinfo"
    If rsCompanyInfo.RecordCount > 0 Then
        lblCompanyInfo.Caption = "This software is licensed to " & UCase(rsCompanyInfo!CompanyName) & "."
    End If
    
    Exit Sub
    
errH:

    If err.Number = -2147467259 Then
        If MsgBox("Database not found. Do you want to configure the Database Parameter?", vbQuestion + vbYesNo) = vbYes Then
            Unload Me
            frmUtilDatabaseConfig.Show
            
        Else
            End
        End If
    Else
        MsgBox "Error code :" & err.Number & vbCrLf & _
            "Error description : " & err.Description, vbCritical, "ERROR!"
    End If

End Sub

Private Sub txtPassword_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtUsername.Text) = "" Then
            txtUsername.SetFocus
        Else
            cmdOK_Click
        End If
    End If
End Sub

Private Sub txtUserName_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtPassword.Text) = "" Then
            SendKeys "{TAB}"
        Else
            cmdOK_Click
        End If
    End If
End Sub
