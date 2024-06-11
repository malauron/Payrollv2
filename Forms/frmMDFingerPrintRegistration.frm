VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmMDFingerPrintRegistration 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Fingerprint Registration"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   0
      ScaleHeight     =   2730
      ScaleWidth      =   2730
      TabIndex        =   2
      Top             =   1215
      Width           =   2760
   End
   Begin VB.ListBox Status 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   2775
      TabIndex        =   1
      Top             =   1215
      Width           =   4395
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   2145
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   3510
      Visible         =   0   'False
      Width           =   615
   End
   Begin TDBText6Ctl.TDBText txtLastname 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   825
      Width           =   2760
      _Version        =   65536
      _ExtentX        =   4868
      _ExtentY        =   529
      Caption         =   "frmMDFingerPrintRegistration.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMDFingerPrintRegistration.frx":006C
      Key             =   "frmMDFingerPrintRegistration.frx":008A
      BackColor       =   16777215
      EditMode        =   0
      ForeColor       =   4210752
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
      MaxLength       =   20
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
   Begin TDBText6Ctl.TDBText txtFirstname 
      Height          =   300
      Left            =   2775
      TabIndex        =   5
      Top             =   825
      Width           =   3330
      _Version        =   65536
      _ExtentX        =   5874
      _ExtentY        =   529
      Caption         =   "frmMDFingerPrintRegistration.frx":00CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMDFingerPrintRegistration.frx":013A
      Key             =   "frmMDFingerPrintRegistration.frx":0158
      BackColor       =   16777215
      EditMode        =   0
      ForeColor       =   4210752
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
      MaxLength       =   20
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
   Begin TDBText6Ctl.TDBText txtMiddleName 
      Height          =   300
      Left            =   6120
      TabIndex        =   6
      Top             =   825
      Width           =   2670
      _Version        =   65536
      _ExtentX        =   4710
      _ExtentY        =   529
      Caption         =   "frmMDFingerPrintRegistration.frx":019C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMDFingerPrintRegistration.frx":0208
      Key             =   "frmMDFingerPrintRegistration.frx":0226
      BackColor       =   16777215
      EditMode        =   0
      ForeColor       =   4210752
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
      MaxLength       =   20
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
   Begin C1SizerLibCtl.C1Elastic cmdSearch 
      Height          =   315
      Left            =   8820
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   810
      Width           =   315
      _cx             =   556
      _cy             =   556
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmMDFingerPrintRegistration.frx":026A
      Caption         =   ":::"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
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
   Begin C1SizerLibCtl.C1Elastic cmdSave 
      Height          =   675
      Left            =   7275
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1665
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmMDFingerPrintRegistration.frx":0286
      Caption         =   "SAVE [F2]"
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
   Begin C1SizerLibCtl.C1Elastic cmdClose 
      Height          =   675
      Left            =   7275
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2670
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmMDFingerPrintRegistration.frx":0F60
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
   Begin TDBText6Ctl.TDBText txtPrompt 
      Height          =   300
      Left            =   15
      TabIndex        =   15
      Top             =   255
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   529
      Caption         =   "frmMDFingerPrintRegistration.frx":1C3A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMDFingerPrintRegistration.frx":1CA6
      Key             =   "frmMDFingerPrintRegistration.frx":1CC4
      BackColor       =   16777215
      EditMode        =   0
      ForeColor       =   4210752
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
      Text            =   "Select an employee."
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fingerprint samples needed:"
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
      Height          =   210
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   4125
      Width           =   2520
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt"
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
      Height          =   210
      Index           =   1
      Left            =   15
      TabIndex        =   11
      Top             =   15
      Width           =   1725
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
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
      Height          =   210
      Index           =   2
      Left            =   5490
      TabIndex        =   9
      Top             =   615
      Width           =   1725
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Height          =   210
      Index           =   1
      Left            =   1965
      TabIndex        =   8
      Top             =   615
      Width           =   1725
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Height          =   210
      Index           =   0
      Left            =   15
      TabIndex        =   7
      Top             =   615
      Width           =   1725
   End
   Begin VB.Label Samples 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2535
      TabIndex        =   3
      Top             =   4065
      Width           =   615
   End
End
Attribute VB_Name = "frmMDFingerPrintRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Capture  As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs          As DPFPFeatureExtraction
Dim CreateTempl         As DPFPEnrollment
Dim ConvertSample       As DPFPSampleConversion
Dim Templ               As DPFPTemplate
Public mEmployeeCode    As Integer

Private Sub DrawPicture(ByVal Pict As IPictureDisp)
 ' Must use hidden PictureBox to easily resize picture.
 Set HiddenPict.Picture = Pict
 Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
 Picture1.Picture = Picture1.Image
End Sub

Private Sub ReportStatus(ByVal str As String)
 ' Add string to list box.
 Status.AddItem (str)
 ' Move list box selection down.
 Status.ListIndex = Status.NewIndex
End Sub

Private Sub cmdClose_Click()
 ' Stop capture operation. This code is optional.
 Capture.StopCapture
 ' Unload form.
 Unload Me
End Sub

Private Sub cmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdClose.Appearance = apInsetLight
  End If
End Sub

Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdClose.Appearance = apRaisedLight
  End If
End Sub

Private Sub cmdSave_Click()

  Dim blob()          As Byte
  
  Dim rsFinger        As ADODB.Recordset
  Dim myFinger        As ADODB.Stream
  
  If mEmployeeCode = 0 Then
    MsgBox "Please select an employee.", vbExclamation + vbOKOnly
    Exit Sub
  End If
  
  ' First verify that template is not empty.
  If Templ Is Nothing Then
   MsgBox "Please create a fingerprint template.", vbExclamation + vbOKOnly
   Exit Sub
  End If
  
  If Dir(App.Path & "\FingerPrintTemplates", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\FingerPrintTemplates")
  End If
   
  blob = Templ.Serialize
  ' Save binary data to file.
  Open App.Path & "\FingerPrintTemplates\" & mEmployeeCode & ".fpt" For Binary As #1
  Put #1, , blob
  Close #1
 
  ConMain.Execute "delete from finger where employeecode = " & mEmployeeCode & ""
  
  Set myFinger = New ADODB.Stream
  myFinger.Type = adTypeBinary
  myFinger.Open
  myFinger.LoadFromFile App.Path & "\FingerPrintTemplates\" & mEmployeeCode & ".fpt"
  
  NetOpen rsFinger, "select * from finger where employeecode = " & mEmployeeCode & ""
  
  With rsFinger
      .AddNew
      .Fields("employeecode") = mEmployeeCode
      .Fields("fingercode") = myFinger.Read
      .Update
  End With
  
  myFinger.Close
 
 MsgBox "Employee fingerprint template was succesfully saved.", vbInformation + vbOKOnly
 
  Reset_Template
  
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdSave.Appearance = apInsetLight
  End If
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdSave.Appearance = apRaisedLight
  End If
End Sub

Private Sub cmdSearch_Click()
  With frmBrowseEmployee
    .mBrowseType = "FingerPrintReg"
    .Show vbModal
  End With
End Sub

Private Sub cmdSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdSearch.Appearance = apInsetLight
  End If
End Sub

Private Sub cmdSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    cmdSearch.Appearance = apRaisedLight
  End If
End Sub

Private Sub Form_Load()
  Retreive_Fingerprints
  Reset_Template
End Sub

Public Sub Reset_Template()
 
 mEmployeeCode = 0
 
 Set Templ = Nothing
 
 txtPrompt.Text = "Select an employee."
 
 Status.Clear
 
 txtLastname.Text = ""
 txtFirstname.Text = ""
 txtMiddleName.Text = ""
 
 Picture1.Picture = Nothing
 ' Create capture operation.
 Set Capture = New DPFPCapture
 ' Start capture operation.
 Capture.StartCapture
 ' Create DPFPFeatureExtraction object.
 Set CreateFtrs = New DPFPFeatureExtraction
 ' Create DPFPEnrollment object.
 Set CreateTempl = New DPFPEnrollment
 ' Show number of samples needed.
 Samples.Caption = CreateTempl.FeaturesNeeded
 ' Create DPFPSampleConversion object.
 Set ConvertSample = New DPFPSampleConversion
 
End Sub

Private Sub Capture_OnReaderConnect(ByVal ReaderSerNum As String)
  ReportStatus ("The fingerprint reader was connected.")
End Sub

Private Sub Capture_OnReaderDisconnect(ByVal ReaderSerNum As String)
 ReportStatus ("The fingerprint reader was disconnected.")
End Sub

Private Sub Capture_OnFingerTouch(ByVal ReaderSerNum As String)
  If mEmployeeCode > 0 Then
    ReportStatus ("The fingerprint reader was touched.")
  Else
    ReportStatus ("You must first select an employee. Operation cancelled.")
  End If
End Sub
Private Sub Capture_OnFingerGone(ByVal ReaderSerNum As String)
  If mEmployeeCode > 0 Then
    ReportStatus ("The finger was removed from the fingerprint reader.")
  End If
End Sub
Private Sub Capture_OnSampleQuality(ByVal ReaderSerNum As String, ByVal Feedback As DPFPCaptureFeedbackEnum)
 If mEmployeeCode > 0 Then
    If Feedback = CaptureFeedbackGood Then
      ReportStatus ("The quality of the fingerprint sample is good.")
    Else
      ReportStatus ("The quality of the fingerprint sample is poor.")
    End If
  End If
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)


 Dim Feedback As DPFPCaptureFeedbackEnum
  
  If mEmployeeCode > 0 Then
    ReportStatus ("The fingerprint sample was captured.")
    ' Draw fingerprint image.
    DrawPicture ConvertSample.ConvertToPicture(Sample)
    ' Process sample and create feature set for purpose of enrollment.
    Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeEnrollment)
    ' Quality of sample is not good enough to produce feature set.
    If Feedback = CaptureFeedbackGood Then
     ReportStatus ("The fingerprint feature set was created.")
     txtPrompt.Text = "Touch the fingerprint reader again with the same finger."
     ' Add feature set to template.
     CreateTempl.AddFeatures CreateFtrs.FeatureSet
     ' Show number of samples needed to complete template.
     Samples.Caption = CreateTempl.FeaturesNeeded
     ' Check if template has been created.
     If CreateTempl.TemplateStatus = TemplateStatusTemplateReady Then
       SetTemplete CreateTempl.Template
       ' Template has been created, so stop capturing samples.
       Capture.StopCapture
       txtPrompt.Text = "Click Save."
       MsgBox "The fingerprint template was created.", vbInformation + vbOKOnly
     End If
    End If
  End If
End Sub

Public Function GetTemplate() As Object
 ' Template can be empty. If so, then returns Nothing.
 If Templ Is Nothing Then
 Else
  Set GetTemplate = Templ
 End If
End Function

Public Sub SetTemplete(ByVal Template As Object)
 Set Templ = Template
End Sub

Private Sub Retreive_Fingerprints()

  Dim rsFinger        As ADODB.Recordset
  Dim myFinger        As ADODB.Stream

  If Dir(App.Path & "\FingerPrintTemplates", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\FingerPrintTemplates")
  End If
  
  NetOpen rsFinger, "select * from finger"
  
  With rsFinger
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
      
        If Dir(App.Path & "\FingerPrintTemplates\" & !employeecode & ".fpt") <> vbNullString Then
          Kill App.Path & "\FingerPrintTemplates\" & !employeecode & ".fpt"
        End If
        
        Set myFinger = New ADODB.Stream
        myFinger.Type = adTypeBinary
        myFinger.Open
        myFinger.Write .Fields("fingercode") 'retrieve the image from the database
        myFinger.SaveToFile App.Path & "\FingerPrintTemplates\" & !employeecode & ".fpt" 'write the retrieved image to the database
        myFinger.Close 'close the created stream
        
        .MoveNext
        
      Loop
    End If
  End With
  
End Sub
