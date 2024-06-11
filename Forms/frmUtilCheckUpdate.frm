VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilCheckUpdate 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check For Updates"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      TabIndex        =   4
      Top             =   510
      Width           =   6225
      Begin lvButton.lvButtons_H cmdClose 
         Cancel          =   -1  'True
         Height          =   390
         Left            =   1215
         TabIndex        =   5
         Top             =   45
         Width           =   1995
         _ExtentX        =   3519
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
         Image           =   "frmUtilCheckUpdate.frx":0000
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   390
         Left            =   3225
         TabIndex        =   6
         Top             =   45
         Width           =   1995
         _ExtentX        =   3519
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
         Image           =   "frmUtilCheckUpdate.frx":0CDA
         cBack           =   14737632
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   540
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   6255
      Begin lvButton.lvButtons_H cmdShowRegistration 
         Height          =   315
         Left            =   5835
         TabIndex        =   1
         ToolTipText     =   "Browse for checked in guests."
         Top             =   165
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Caption         =   "..."
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
         cBack           =   14737632
      End
      Begin TDBText6Ctl.TDBText txtPath 
         Height          =   300
         Left            =   525
         TabIndex        =   2
         Top             =   165
         Width           =   5280
         _Version        =   65536
         _ExtentX        =   9313
         _ExtentY        =   529
         Caption         =   "frmUtilCheckUpdate.frx":19B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmUtilCheckUpdate.frx":1A20
         Key             =   "frmUtilCheckUpdate.frx":1A3E
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   4210752
         ReadOnly        =   0
         ShowContextMenu =   1
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
         ScrollBars      =   2
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
         HighlightText   =   -1
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   -2205
         TabIndex        =   3
         Top             =   210
         Width           =   2670
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   510
      Top             =   1545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.exe"
   End
End
Attribute VB_Name = "frmUtilCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mOldApp                         As String

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdOK_Click()

    On Error GoTo Err_Hndlr:
    
    
    Dim Phrase              As String
    Dim mString             As String
    Dim Char1               As String
    Dim FileName1           As String
    
    Dim Position            As Integer
    Dim ctr                 As Integer
    
    Dim Asc1                As Long
    
        
    If DirExists(txtPath.Text) = True And Trim(txtPath.Text) <> "" Then

Err_Retry:
        
        FileCopy txtPath.Text, App.Path & "\" & mOldApp
        
        
        
        Phrase = txtPath.Text & "," & 0 & "," & mOldApp & ","
        mString = ""
        
        For Position = Len(Phrase) To 1 Step -1
            Char1 = Mid$(Phrase, Position, 1)
            Asc1 = Asc(Char1)
            Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
            Char1 = Chr$(Asc1)
            mString = mString & Char1
        Next
        
        FileName1 = App.Path & "\Extdata\upd8.dat"
        
        Open FileName1 For Output As #1
            Print #1, mString
        Close #1
    
        MsgBox "Restarting program.", vbInformation + vbOKOnly
        
        Shell mOldApp
        
        End

    Else
    
        MsgBox "Path not found.", vbExclamation + vbOKOnly
        
    End If
    
    Exit Sub
    
Err_Hndlr:

    If MsgBox(Err.Description & vbCrLf & "Payroll system is still on the termination process. Please try again.", vbCritical + vbOKCancel) = vbCancel Then
        End
    End If
    
End Sub

Private Sub cmdShowRegistration_Click()
        
    dlg.FileName = "*.exe"
    dlg.ShowOpen
    If Len(dlg.FileName) = 0 Then Exit Sub

    MousePointer = MousePointerConstants.vbHourglass
    MousePointer = MousePointerConstants.vbDefault
    
    txtPath.Text = dlg.FileName

End Sub

Private Sub Form_Load()

    Dim FileName1       As String
    Dim txt             As String
    Dim Phrase          As String
    Dim Char1           As String
    Dim mEncrypted      As String
    Dim mSourcePath     As String
    Dim mString         As String
    
    Dim Position        As Integer
    Dim ctr             As Integer
    
    Dim Asc1            As Long
    
    Dim mToUpdate       As Boolean
    
    FileName1 = App.Path & "\ExtData\upd8.dat"
    
    mEncrypted = ""
    
    Open FileName1 For Input As #1
        Do Until EOF(1)
            Input #1, txt
            mEncrypted = mEncrypted & txt
        Loop
    Close #1
    
    Phrase = mEncrypted
    
    mEncrypted = ""
    
    ctr = 0
    
    mToUpdate = False
    
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        Asc1 = Asc(Char1)
        Asc1 = (((Asc1 * Asc1) / 2) / 2)
        Asc1 = Sqr(Asc1)
        Char1 = Chr$(Asc1)
        '----
        If Char1 = "," Then
            Char1 = ""
            ctr = ctr + 1
            If ctr = 1 Then
                mSourcePath = mEncrypted
                mEncrypted = ""
            ElseIf ctr = 2 Then
                If mEncrypted = 1 Then
                    mToUpdate = True
                End If
                mEncrypted = ""
            ElseIf ctr = 3 Then
                mOldApp = mEncrypted
                mEncrypted = ""
            End If
        End If
        '-----
        mEncrypted = mEncrypted & Char1
    Next
    
    If mToUpdate = False Then End
    
    Phrase = mSourcePath & "," & 0 & "," & mOldApp & ","
    mString = ""
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        Asc1 = Asc(Char1)
        Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
        Char1 = Chr$(Asc1)
        mString = mString & Char1
    Next
    
    FileName1 = App.Path & "\Extdata\upd8.dat"
    
    Open FileName1 For Output As #1
        Print #1, mString
    Close #1
    
    txtPath.Text = mSourcePath
    
End Sub

Public Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = False
    ' test the directory attribute
    If Dir(DirName) <> "" Then
        DirExists = True
    End If
    Exit Function
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

