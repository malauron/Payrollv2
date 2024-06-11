VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLeaveCredits 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   8235
   WindowState     =   2  'Maximized
   Begin TDBNumber6Ctl.TDBNumber txtAmount 
      Height          =   270
      Left            =   2595
      TabIndex        =   4
      Top             =   8130
      Visible         =   0   'False
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   476
      Calculator      =   "frmLeaveCredits.frx":0000
      Caption         =   "frmLeaveCredits.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLeaveCredits.frx":008C
      Keys            =   "frmLeaveCredits.frx":00AA
      Spin            =   "frmLeaveCredits.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   1
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,###,##0.00"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1145962501
      MinValueVT      =   1414463493
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      Height          =   765
      Left            =   150
      TabIndex        =   0
      Top             =   570
      Width           =   7875
      Begin TDBNumber6Ctl.TDBNumber txtPayyear 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   270
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   556
         Calculator      =   "frmLeaveCredits.frx":011C
         Caption         =   "frmLeaveCredits.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmLeaveCredits.frx":01A2
         Keys            =   "frmLeaveCredits.frx":01C0
         Spin            =   "frmLeaveCredits.frx":020A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#####"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1145962501
         MinValueVT      =   1414463493
      End
      Begin lvButton.lvButtons_H cmdGenerate 
         Height          =   330
         Left            =   3885
         TabIndex        =   5
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         Caption         =   "&Generate"
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
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdSave 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   5475
         TabIndex        =   6
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         Caption         =   "&Save"
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
         cFore           =   4210752
         cFHover         =   4210752
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   4
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year"
         Height          =   255
         Left            =   105
         TabIndex        =   2
         Top             =   345
         Width           =   1560
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgLeaveCredits 
      Height          =   6720
      Left            =   300
      TabIndex        =   3
      Top             =   990
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   11853
      _LayoutType     =   4
      _RowHeight      =   16
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Emp ID"
      Columns(0).DataField=   "empno"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Employee Name"
      Columns(1).DataField=   "fullname"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   4
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4736"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4657"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Verdana"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HF6F8F8&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H400000&"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(17)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFF0EA&"
      _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HF6F8F8&"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=0"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin LinkProPayroll.b8ChildTitleBar TitleBar 
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   609
      Caption         =   "Title"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
   End
End
Attribute VB_Name = "frmLeaveCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click()
  
    Dim rsEmployee    As ADODB.Recordset
    Dim rsLeave       As ADODB.Recordset
    Dim rsEmpLvTmp    As ADODB.Recordset
    Dim rsEmpLvCrdt   As ADODB.Recordset
    
    Dim C             As TrueOleDBGrid80.Column
    Dim I             As Integer
    
    If Not IsNumeric(txtPayyear.Text) Or CDbl(txtPayyear.Text) <= 0 Then
        MsgBox "Please provide a financial year.", vbExclamation + vbOKOnly
        txtPayyear.SetFocus
        Exit Sub
    End If
    
    NetOpen rsEmployee, "select employeecode, concat(lastname,', ',firstname,'',middlename) fullname from employee order by concat(lastname,', ',firstname,'',middlename)"
    
    NetOpen rsLeave, "select leavetypescode, leavetypes from leavetypes order by leavetypescode"
    
    If rsEmployee.RecordCount > 0 Then
        If rsLeave.RecordCount > 0 Then
        
            Set rsEmpLvTmp = New ADODB.Recordset
            
            If tdgLeaveCredits.Columns.count > 2 Then
              Do While tdgLeaveCredits.Columns.count > 2
                tdgLeaveCredits.Columns.Remove (tdgLeaveCredits.Columns.count - 1)
              Loop
            End If
            
            With rsEmpLvTmp
              .Fields.Append "employeecode", adVarChar, 15
              .Fields.Append "fullname", adVarChar, 100
              rsLeave.MoveFirst
              I = 2
              Do While Not rsLeave.EOF
                
                .Fields.Append rsLeave!leavetypescode, adVarChar, 7
                
                Set C = tdgLeaveCredits.Columns.Add(I)
                C.Visible = True
                C.DataField = CStr(rsLeave!leavetypescode)
                C.Caption = CStr(rsLeave!LeaveTypes)
                C.ExternalEditor = "txtAmount"
                C.Alignment = dbgRight
                C.NumberFormat = "#,##0.00"
                I = I + 1
                rsLeave.MoveNext
                DoEvents
              Loop
              .Open
              Set tdgLeaveCredits.DataSource = rsEmpLvTmp
            End With
            
            With rsEmployee
              .MoveFirst
              Do While Not .EOF
                rsEmpLvTmp.AddNew
                rsEmpLvTmp!employeecode = !employeecode
                rsEmpLvTmp!Fullname = !Fullname
                rsEmpLvTmp.Update
                NetOpen rsEmpLvCrdt, "select * from emplvcredits where employeecode = '" & !employeecode & "' and payyear = '" & txtPayyear.Text & "'"
                If rsEmpLvCrdt.RecordCount > 0 Then
                  I = 2
                  For I = 2 To tdgLeaveCredits.Columns.count - 1
                    rsEmpLvCrdt.MoveFirst
                    rsEmpLvCrdt.Find "leavetypescode = '" & tdgLeaveCredits.Columns(I).DataField & "'"
                    If Not rsEmpLvCrdt.EOF Then
                      If rsEmpLvCrdt!nooflv > 0 Then
                        tdgLeaveCredits.Columns(I).Text = rsEmpLvCrdt!nooflv
                      End If
                    End If
                  Next
                End If
                rsEmpLvTmp.Update
                .MoveNext
                DoEvents
              Loop
              rsEmpLvTmp.MoveFirst
            End With
            
            If rsEmpLvTmp.RecordCount > 0 Then
              cmdSave.Enabled = True
            Else
              cmdSave.Enabled = False
            End If
            
        Else
        
            cmdSave.Enabled = False
        
        End If
        
    Else
    
        cmdSave.Enabled = False
    
    End If

  
End Sub

Private Sub cmdSave_Click()

  Dim I             As Integer
  
  With tdgLeaveCredits
    If Not .EOF Then
      .MoveFirst
      
      ConMain.Execute "set autocommit = 0"
      ConMain.BeginTrans

      Do While Not .EOF
        ConMain.Execute "delete from emplvcredits where employeecode = '" & .Columns("employeecode").Text & "' and payyear = '" & txtPayyear.Text & "'"
        For I = 2 To .Columns.count - 1
          If IsNumeric(.Columns(I).Text) Then
            If CDbl(.Columns(I).Text) > 0 Then
              ConMain.Execute "insert into emplvcredits (employeecode,payyear,leavetypescode,nooflv) values " & _
                                    "('" & .Columns("employeecode").Text & "','" & txtPayyear.Text & "', " & _
                                    "'" & .Columns(I).DataField & "','" & .Columns(I).Text & "')"
            End If
          End If
        Next
      
        .MoveNext
      Loop
      ConMain.CommitTrans
      .MoveLast
    End If
  End With
End Sub

Private Sub Form_Load()
    
    Add_MDIButton Me.Name, titlebar.Caption
    
    cmdSave.Enabled = False
    txtPayyear = Format(Now, "YYYY")
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Remove_MDIButton Me.Name
End Sub

Private Sub Form_Resize()

    titlebar.Move 0, 0, Me.ScaleWidth

  With fra1
    .Top = titlebar.Top + titlebar.Height
    .Left = 0
    .Width = Me.ScaleWidth
  End With
  
  With tdgLeaveCredits
    .Top = fra1.Height + fra1.Top
    .Left = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight - fra1.Height
  End With
  
End Sub
