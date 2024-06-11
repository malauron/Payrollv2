VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilImportTito 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   9270
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsOldTITO 
      Height          =   4095
      Left            =   285
      TabIndex        =   0
      Top             =   1260
      Width           =   5625
      _cx             =   9922
      _cy             =   7223
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16185592
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16185592
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin CitronSoftwarePayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Import TITO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   3186872
      GradTheme       =   2
   End
   Begin CitronSoftwarePayroll.b8SContainer frabutton 
      Height          =   585
      Left            =   45
      TabIndex        =   2
      Top             =   525
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      BorderColor     =   14215660
      Begin lvButton.lvButtons_H cmdImport 
         Height          =   420
         Left            =   75
         TabIndex        =   3
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Import"
         CapAlign        =   2
         BackStyle       =   5
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
         cBhover         =   14215660
         Focus           =   0   'False
         cGradient       =   14215660
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   16185592
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsNewTITO 
      Height          =   4095
      Left            =   210
      TabIndex        =   4
      Top             =   5520
      Width           =   5625
      _cx             =   9922
      _cy             =   7223
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16185592
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16185592
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   105
      Left            =   0
      TabIndex        =   5
      Top             =   1095
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmUtilImportTito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsOldTito        As ADODB.Recordset
Dim rsNewTito        As ADODB.Recordset

Private Sub cmdImport_Click()

    Dim mEmpNo      As String
    Dim mCompLog    As String
    
    Dim rsCheck     As ADODB.Recordset
    
    
    With rsOldTito
        If .RecordCount > 0 Then
        
            pb1.Max = .RecordCount
            pb1.Value = 0
            .MoveFirst
            
            CitronPayroll.Execute "set autocommit = 0"
            CitronPayroll.BeginTrans
            Do While Not .EOF
            
                pb1.Value = pb1.Value + 1
                
                mEmpNo = "0" & !debt_code
                
                mCompLog = Format(!tito_date & " " & !tito_time, "YYYY-MM-DD hh:nn:ss")
                
                CitronPayroll.Execute "insert into tito(empno,biometid,complog,datelog,timelog,logstat) values " & _
                    "('" & mEmpNo & "','" & mEmpNo & "','" & mCompLog & "', " & _
                    "'" & Format(!tito_date, "YYYY-MM-DD") & "','" & Format(!tito_time, "hh:nn:ss") & "','" & IIf(Trim(!Type) = "O", "Out", "In") & "')"
                
                .MoveNext
                
                DoEvents
                
            Loop
            
            CitronPayroll.CommitTrans
            
            MsgBox "Download complete.", vbExclamation + vbOKOnly
            
            rsNewTito.Requery
            
        Else
        
            MsgBox "Tax Table table is empty.", vbExclamation + vbOKOnly
            Exit Sub
            
        End If
    End With
    
End Sub

Private Sub Form_Load()

    Add_MDIButton Me.Name, TitleBar.Caption
    
    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    Set rsOldTito = New ADODB.Recordset
    rsOldTito.Open "select * from tito", ConAdvPayroll, adOpenStatic, adLockOptimistic
    Set vsOldTITO.DataSource = rsOldTito
        
    NetOpen rsNewTito, "", "select * from tito"
    Set vsNewTITO.DataSource = rsNewTito
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Remove_MDIButton Me.Name
    
End Sub

Private Sub Form_Resize()
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
        .Top = TitleBar.Top + TitleBar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With pb1
        .Top = frabutton.Top + frabutton.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With vsOldTITO
        .Top = pb1.Top + pb1.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = (Me.ScaleHeight - (pb1.Top + pb1.Height)) / 2
    End With
    
    With vsNewTITO
        .Top = vsOldTITO.Top + vsOldTITO.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = (Me.ScaleHeight - (pb1.Top + pb1.Height)) / 2
    End With
    
End Sub




