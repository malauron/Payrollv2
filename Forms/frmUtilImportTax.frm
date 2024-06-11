VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUtilImportTax 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   6510
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsOldTax 
      Height          =   4095
      Left            =   105
      TabIndex        =   0
      Top             =   1275
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
      Caption         =   "Import Tax Table"
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
   Begin VSFlex8Ctl.VSFlexGrid vsNewTax 
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
Attribute VB_Name = "frmUtilImportTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsOldTax        As ADODB.Recordset
Dim rsNewTax        As ADODB.Recordset

Private Sub cmdImport_Click()

    Dim mWTCode     As String
    
    With rsOldTax
        If .RecordCount > 0 Then
        
            pb1.Max = .RecordCount
            pb1.Value = 0
            .MoveFirst
            
            CitronPayroll.Execute "set autocommit = 0"
            CitronPayroll.BeginTrans
            Do While Not .EOF
            
                pb1.Value = pb1.Value + 1
                mWTCode = LastCode("GetLastCodeA", "WT", "0000000")
                
                CitronPayroll.Execute "insert into wt(wtcode,description,exemption," & _
                    "b1,b2,b3,b4,b5,b6,b7,b8,b9,b10," & _
                    "f1,f2,f3,f4,f5,f6,f7,f8,f9,f10," & _
                    "a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,dummycode) " & _
                    "values ('" & mWTCode & "', '" & !brckt_desc & "', " & !exemption & ", " & _
                    "" & !brckt_01 & "," & !brckt_02 & "," & !brckt_03 & "," & !brckt_04 & "," & !brckt_05 & "," & !brckt_06 & "," & !brckt_07 & "," & !brckt_08 & "," & !brckt_09 & "," & !brckt_10 & ", " & _
                    "" & !factor1 & "," & !factor2 & "," & !factor3 & "," & !factor4 & "," & !factor5 & "," & !factor6 & "," & !factor7 & "," & !factor8 & "," & !factor9 & "," & !factor10 & ", " & _
                    "" & !add_on1 & "," & !add_on2 & "," & !add_on3 & "," & !add_on4 & "," & !add_on5 & "," & !add_on6 & "," & !add_on7 & "," & !add_on8 & "," & !add_on9 & "," & !add_on10 & ",'" & !Brckt_No & "')"
                    
                .MoveNext
                DoEvents
            Loop
            
            CitronPayroll.CommitTrans
            
            rsNewTax.Requery
            
        Else
        
            MsgBox "Tax Table table is empty.", vbExclamation + vbOKOnly
            Exit Sub
            
        End If
    End With
    
End Sub

Private Sub Form_Load()

    Add_MDIButton Me.Name, titlebar.Caption
    
    SendMessage pb1.hwnd, &H400 + 9, 0, RGB(99, 138, 231)
    SendMessage pb1.hwnd, &H2000 + 1, 0, RGB(255, 255, 255)
    
    Set rsOldTax = New ADODB.Recordset
    rsOldTax.Open "select * from whtax", ConAdvPayroll, adOpenStatic, adLockOptimistic
    Set vsOldTax.DataSource = rsOldTax
        
    NetOpen rsNewTax, "", "select * from wt"
    Set vsNewTax.DataSource = rsNewTax
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Remove_MDIButton Me.Name
    
End Sub

Private Sub Form_Resize()
    
    titlebar.Move 0, 0, Me.ScaleWidth
    
    With frabutton
        .Top = titlebar.Top + titlebar.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With pb1
        .Top = frabutton.Top + frabutton.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With vsOldTax
        .Top = pb1.Top + pb1.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = (Me.ScaleHeight - (pb1.Top + pb1.Height)) / 2
    End With
    
    With vsNewTax
        .Top = vsOldTax.Top + vsOldTax.Height
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = (Me.ScaleHeight - (pb1.Top + pb1.Height)) / 2
    End With
    
End Sub


