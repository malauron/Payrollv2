VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilCompanyInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Information"
   ClientHeight    =   6300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmVATTypes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "VAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3015
      Left            =   1215
      TabIndex        =   6
      Top             =   2535
      Width           =   4620
      Begin VB.Frame frmVATRates 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "VAT Rate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2025
         Left            =   375
         TabIndex        =   9
         Top             =   885
         Width           =   4125
         Begin TDBNumber6Ctl.TDBNumber txtVATRate 
            Height          =   300
            Left            =   1455
            TabIndex        =   15
            Top             =   600
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   529
            Calculator      =   "frmUtilCompanyInfo.frx":0000
            Caption         =   "frmUtilCompanyInfo.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmUtilCompanyInfo.frx":0086
            Keys            =   "frmUtilCompanyInfo.frx":00A4
            Spin            =   "frmUtilCompanyInfo.frx":00EE
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##0.00%"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   4194304
            Format          =   "##0.00%"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   100
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   2088828933
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.OptionButton optVariable 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Variable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   435
            TabIndex        =   12
            Top             =   615
            Width           =   1545
         End
         Begin VB.OptionButton optZeroRated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Zero Rated"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   435
            TabIndex        =   11
            Top             =   300
            Width           =   1545
         End
         Begin VB.Frame frmType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   990
            Left            =   645
            TabIndex        =   10
            Top             =   930
            Width           =   3375
            Begin VB.OptionButton optInclusive 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Inclusive"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   615
               Width           =   1545
            End
            Begin VB.OptionButton optExclusive 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Exclusive"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   300
               Width           =   1545
            End
         End
      End
      Begin VB.OptionButton optVAT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "VAT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   555
         Width           =   1545
      End
      Begin VB.OptionButton optNonVat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Non-VAT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   1545
      End
   End
   Begin TDBText6Ctl.TDBText txtCompanyName 
      Height          =   615
      Left            =   1230
      TabIndex        =   0
      Top             =   390
      Width           =   4620
      _Version        =   65536
      _ExtentX        =   8149
      _ExtentY        =   1085
      Caption         =   "frmUtilCompanyInfo.frx":0116
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilCompanyInfo.frx":0182
      Key             =   "frmUtilCompanyInfo.frx":01A0
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
      AlignVertical   =   0
      MultiLine       =   -1
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
   Begin TDBText6Ctl.TDBText txtCompanyAddress 
      Height          =   1455
      Left            =   1230
      TabIndex        =   1
      Top             =   1050
      Width           =   4620
      _Version        =   65536
      _ExtentX        =   8149
      _ExtentY        =   2566
      Caption         =   "frmUtilCompanyInfo.frx":01E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmUtilCompanyInfo.frx":0250
      Key             =   "frmUtilCompanyInfo.frx":026E
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
      AlignVertical   =   0
      MultiLine       =   -1
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
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   5805
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      Caption         =   "&Update"
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
      Image           =   "frmUtilCompanyInfo.frx":02B2
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H cmdClose 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   3345
      TabIndex        =   3
      Top             =   5805
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      Caption         =   "E&xit"
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
      Image           =   "frmUtilCompanyInfo.frx":0B8C
      cBack           =   14737632
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Address"
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
      Height          =   780
      Left            =   315
      TabIndex        =   5
      Top             =   1515
      Width           =   1365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
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
      Height          =   570
      Left            =   330
      TabIndex        =   4
      Top             =   510
      Width           =   1095
   End
End
Attribute VB_Name = "frmUtilCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCompanyInfo As ADODB.Recordset
    

Private Sub Form_Load()
    
    Set rsCompanyInfo = New ADODB.Recordset
    
    NetOpen rsCompanyInfo, "select * from companyinfo"
    
    With rsCompanyInfo
        If .RecordCount > 0 Then
            txtCompanyName.Text = !CompanyName
            txtCompanyAddress.Text = !companyaddress
            optNonVat.Value = IIf(!nonvat = "T", True, False)
            optVAT.Value = IIf(!vat = "T", True, False)
            optZeroRated.Value = IIf(!zerorated = "T", True, False)
            optVariable.Value = IIf(!variable = "T", True, False)
            optExclusive.Value = IIf(!Exclusive = "T", True, False)
            optInclusive.Value = IIf(!inclusive = "T", True, False)
            txtVATRate.Text = Format(!vatrate, "##0.00%")
        End If
    End With
    
End Sub

Private Sub cmdUpdate_Click()

    If Trim(txtCompanyName.Text) = "" Then
        MsgBox "Please enter the name of the company.", vbExclamation + vbOKOnly
        txtCompanyName.SetFocus
        Exit Sub
    End If
    If Trim(txtCompanyAddress.Text) = "" Then
        MsgBox "Please enter the address of the company.", vbExclamation + vbOKOnly
        txtCompanyAddress.SetFocus
        Exit Sub
    End If
    
    ConMain.Execute "set autocommit = 0 "
    ConMain.BeginTrans
    If rsCompanyInfo.RecordCount < 1 Then
        ConMain.Execute "insert into companyinfo (companyname,companyaddress,nonvat,vat,zerorated, " & _
                        "variable,exclusive,inclusive,vatrate) values ('" & Swap(txtCompanyName.Text) & "','" & Swap(txtCompanyAddress.Text) & "','" & IIf(optNonVat.Value = True, "T", "F") & "','" & IIf(optVAT.Value = True, "T", "F") & "','" & IIf(optZeroRated.Value = True, "T", "F") & "', " & _
                        "'" & IIf(optVariable.Value = True, "T", "F") & "','" & IIf(optExclusive.Value = True, "T", "F") & "','" & IIf(optInclusive.Value = True, "T", "F") & "'," & Format(txtVATRate.Text, "##0.00") & ")"
    Else
        ConMain.Execute "update companyinfo set companyname = '" & Swap(txtCompanyName.Text) & "', companyaddress = '" & Swap(txtCompanyAddress.Text) & "',nonvat = '" & IIf(optNonVat.Value = True, "T", "F") & "',vat = '" & IIf(optVAT.Value = True, "T", "F") & "',zerorated = '" & IIf(optZeroRated.Value = True, "T", "F") & "', " & _
                        "variable = '" & IIf(optVariable.Value = True, "T", "F") & "',exclusive = '" & IIf(optExclusive.Value = True, "T", "F") & "',inclusive = '" & IIf(optInclusive.Value = True, "T", "F") & "',vatrate = " & Format(txtVATRate.Text, "##0.00") & ""
    End If
    ConMain.CommitTrans
    
    rsCompanyInfo.Requery
    
    MsgBox "Company information was succesfully updated.", vbInformation + vbOKOnly
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub txtCompanyName_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCompanyAddress_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

    
