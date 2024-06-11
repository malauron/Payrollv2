VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmTimeLog 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.Frame fra3 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   4710
      TabIndex        =   7
      Top             =   2985
      Width           =   5475
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lastname, Firstname Middlename"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1305
         Left            =   -1950
         TabIndex        =   8
         Top             =   135
         Width           =   10425
      End
   End
   Begin VB.Frame fra2 
      BackColor       =   &H00FFFFFF&
      Height          =   6000
      Left            =   15
      TabIndex        =   3
      Top             =   2865
      Width           =   4410
      Begin TDBText6Ctl.TDBText txtIDnumber 
         Height          =   390
         Left            =   105
         TabIndex        =   4
         Top             =   5250
         Width           =   4185
         _Version        =   65536
         _ExtentX        =   7382
         _ExtentY        =   688
         Caption         =   "frmTimeLog.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmTimeLog.frx":006C
         Key             =   "frmTimeLog.frx":008A
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
         Appearance      =   2
         BorderStyle     =   1
         AlignHorizontal =   2
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   5
         Top             =   5610
         Width           =   4035
      End
      Begin VB.Image imgPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   5100
         Left            =   120
         Stretch         =   -1  'True
         Top             =   135
         Width           =   4170
      End
   End
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      Height          =   2325
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   10650
      Begin VB.Timer tmr1 
         Interval        =   1
         Left            =   75
         Top             =   195
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12:00:00 AM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   1305
         Left            =   150
         TabIndex        =   6
         Top             =   1020
         Width           =   10425
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Septembder 31, 2008 Thursday"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   810
         Left            =   135
         TabIndex        =   2
         Top             =   165
         Width           =   10425
      End
   End
   Begin TrueOleDBGrid80.TDBGrid tdgTITO 
      Height          =   4245
      Left            =   4785
      TabIndex        =   0
      Top             =   4020
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   7488
      _LayoutType     =   4
      _RowHeight      =   27
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Log Type"
      Columns(0).DataField=   "logstat"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Time Log"
      Columns(1).DataField=   "complog"
      Columns(1).NumberFormat=   "hh:nn:ss am/pm"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4101"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3889"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3757"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3545"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=18,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=18,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   16777215
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=-1,.fontsize=1800"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&H800000&"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1800,.italic=0"
      _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=3,.fontname=Arial"
      _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H6FE0FD&"
      _StyleDefs(18)  =   ":id=8,.fgcolor=&H0&"
      _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFF0EA&"
      _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=33"
      _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(43)  =   "Named:id=33:Normal"
      _StyleDefs(44)  =   ":id=33,.parent=0"
      _StyleDefs(45)  =   "Named:id=34:Heading"
      _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   ":id=34,.wraptext=-1"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=37:Caption"
      _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(54)  =   "Named:id=38:HighlightRow"
      _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&HFAE4AB&,.fgcolor=&H0&"
      _StyleDefs(56)  =   "Named:id=39:EvenRow"
      _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(58)  =   "Named:id=40:OddRow"
      _StyleDefs(59)  =   ":id=40,.parent=33"
      _StyleDefs(60)  =   "Named:id=41:RecordSelector"
      _StyleDefs(61)  =   ":id=41,.parent=34"
      _StyleDefs(62)  =   "Named:id=42:FilterBar"
      _StyleDefs(63)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmTimeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTito           As ADODB.Recordset
Dim SecCtr           As Integer
Dim rsEmployee       As ADODB.Recordset

Private Sub form_load()
    NetOpen rsTito, "", "select * from tito where biometid = '' limit 0"
    Set tdgTito.DataSource = rsTito
    imgPhoto.Picture = LoadPicture(App.Path & "\Images\nopic.jpg")
End Sub

Private Sub Form_Resize()
    
    With fra1
      .Top = 0
      .Left = 0
      .Width = Me.ScaleWidth
    End With

    With fra2
      .Top = fra1.Height
      .Left = 0
    End With

    With fra3
      .Top = fra1.Top + fra1.Height
      .Left = fra2.Left + fra2.Width
      .Width = Me.ScaleWidth - fra2.Width
    End With

    With tdgTito
      .Top = fra3.Top + fra3.Height
      .Left = fra2.Width
      .Height = Me.ScaleHeight - (fra1.Height + fra3.Height)
      .Width = Me.ScaleWidth - fra2.Width
      .Columns(0).Width = .Width * 0.5
    End With

    With lblDate
      .Top = 300
      .Left = 0
      .Width = fra1.Width
    End With

    With lblTime
      .Top = 300 + lblDate.Height
      .Left = 0
      .Width = fra1.Width
    End With

    With lblName
      .Top = 150
      .Left = 0
      .Width = fra3.Width
    End With

End Sub

Private Sub tmr1_Timer()
  lblDate.Caption = Format(Now, "MMMM DD,YYYY DDDD")
  lblTime.Caption = Format(Now, "hh:nn:ss AM/PM")
End Sub

Private Sub txtIDnumber_Keypress(KeyAscii As Integer)
  
'  On Error GoTo ErrHnldr
  
  Dim rsChk         As ADODB.Recordset
  Dim mStat         As String
  
  If KeyAscii = 13 Then
  
    If Trim(txtIDnumber.Text) <> "" Then
    
      NetOpen rsChk, "", "select * from employee where biometid = '" & txtIDnumber.Text & "'"

      If rsChk.RecordCount > 0 Then
      
        If rsChk.RecordCount > 0 Then
          If rsChk!logstat = "In" Then
            mStat = "Out"
          Else
            mStat = "In"
          End If
        Else
          mStat = "In"
        End If
        
        CitronPayroll.Execute "set autocommit = 0"
        CitronPayroll.BeginTrans
        CitronPayroll.Execute "insert into tito(empno,biometid,complog,datelog,timelog,logstat,cancelled,remarks) values " & _
                              "('" & rsChk!empno & "','" & txtIDnumber.Text & "','" & Format(Now, "YYYY-MM-DD hh:nn:ss") & "', " & _
                              "'" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "hh:nn") & "','" & mStat & "','N','')"
        CitronPayroll.Execute "update employee set logstat = '" & mStat & "' where biometid = '" & txtIDnumber.Text & "'"
        CitronPayroll.CommitTrans
        
        NetOpen rsTito, "", "select * from tito where datelog = '" & Format(Now, "YYYY-MM-DD") & "' and biometid = '" & txtIDnumber.Text & "' "
        rsTito.Sort = "complog"
        rsTito.MoveLast
        
        Set tdgTito.DataSource = rsTito
        imgPhoto.Picture = LoadPicture(App.Path & "\Images\" & txtIDnumber.Text & ".jpg")
        lblName.Caption = UCase(rsChk!lastname & ", " & " " & rsChk!firstname & " " & rsChk!middlename)
      Else
        Set rsTito = Nothing
        Set tdgTito.DataSource = rsTito
        imgPhoto.Picture = LoadPicture(App.Path & "\Images\NoPic.jpg")
        MsgBox "ID number not found.", vbExclamation + vbOKOnly
        txtIDnumber.SetFocus
      End If
    End If
  End If
  
  Exit Sub

ErrHnldr:
  
  imgPhoto.Picture = LoadPicture(App.Path & "\Images\NoPic.jpg")
  Set tdgTito.DataSource = rsTito
  MsgBox err.Description
  
End Sub


