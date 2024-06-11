VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{B49E66FF-6927-4378-9685-937F14679ADD}#1.0#0"; "DPFPCtlX.dll"
Begin VB.Form frmMDFingerPrintVerification 
   Caption         =   "Fingerprint Verification"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic cmdClose 
      Height          =   675
      Left            =   2745
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1875
      _cx             =   3307
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmMDFingerPrintVerification.frx":0000
      Caption         =   "CLOSE [F12]"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   0
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   2
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin DPFPCtlXLibCtl.DPFPVerificationControl DPFPVerificationControl1 
      Height          =   735
      Left            =   3735
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   195
      Width           =   735
      _cx             =   1296
      _cy             =   1296
      ReaderSerialNumber=   "{00000000-0000-0000-0000-000000000000}"
      Active          =   -1  'True
   End
   Begin VB.Label lblName 
      Height          =   435
      Left            =   150
      TabIndex        =   3
      Top             =   1440
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "To verify your identity, touch fingerprint reader with any enrolled finger."
      Height          =   495
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmMDFingerPrintVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ver                   As DPFPVerification
Dim Templ()               As DPFPTemplate
Dim mEmpno()              As String
Dim CTR                   As Integer

Private Sub ReadTemplate(strPath As String, mArrayNo As Integer)
 
 Dim blob()               As Byte
 
 ' Read binary data from file.
 Open strPath For Binary As #1
 ReDim blob(LOF(1))
 Get #1, , blob()
 Close #1
 
 ' Template can be empty, it must be created first.
 If Templ(mArrayNo) Is Nothing Then Set Templ(mArrayNo) = New DPFPTemplate
 
 ' Import binary data to template.
 Templ(mArrayNo).Deserialize blob
 
End Sub

Private Sub GetTemplates()

    Dim FS              As New FileSystemObject
    Dim FSfolder        As Folder
    Dim mfile           As File
    Dim i               As Integer
    
    Set FSfolder = FS.GetFolder(App.Path & "\FingerPrintTemplates")
    
    For Each mfile In FSfolder.Files
        
        DoEvents
        
        If mfile.Type = "FPT File" Then
          i = i + 1
          ReDim Preserve mEmpno(i) As String
          ReDim Preserve Templ(i) As DPFPTemplate
          
          mEmpno(i - 1) = Mid(mfile.Name, 1, Len(mfile.Name) - 4)
          ReadTemplate CStr(mfile), i - 1
        
          
        End If
    Next mfile
    CTR = i
    Set FSfolder = Nothing
    
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 ' Create DPFPVerification object.
 Set Ver = New DPFPVerification
 GetTemplates
End Sub

Private Sub DPFPVerificationControl1_OnComplete(ByVal Ftrs As Object, ByVal Stat As Object)
 
 Dim rsEmployee       As ADODB.Recordset
 
 Dim Res              As Object
 
 Dim i                As Integer
 
 If CTR = 0 Then
  Stat.Status = EventHandlerStatusFailure
  MsgBox "No fingerprint enrolled in the system.", vbExclamation + vbOKOnly
  Exit Sub
 End If
 
 ' Compare feature set with all stored templates.
 For i = LBound(mEmpno) To UBound(mEmpno) - 1
      ' Compare feature set with particular template.
      Set Res = Ver.Verify(Ftrs, Templ(i))
      ' If match, exit from loop.
      If Res.Verified = True Then Exit For
 Next
 
 If Res Is Nothing Then
  Stat.Status = EventHandlerStatusFailure
  lblName.Caption = "Fingerprint was not verified."
  Exit Sub
 ElseIf Res.Verified = False Then
  ' If non-match, notify caller.
  Stat.Status = EventHandlerStatusFailure
  lblName.Caption = "Fingerprint was not verified."
 Else
  NetOpen rsEmployee, "select concat(lastname,', ',firstname,' ',middlename) fullname from employee where employeecode = " & mEmpno(i) & ""
  lblName.Caption = "Fingerprint verified! " & vbCrLf & "Fingerprint belongs to " & rsEmployee!fullname
 End If
 
 
 End Sub
