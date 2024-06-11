VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSysSignatory 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra1 
      BackColor       =   &H00F6F8F8&
      Height          =   6285
      Left            =   15
      TabIndex        =   8
      Top             =   840
      Width           =   6960
      Begin VB.Frame Frame1 
         BackColor       =   &H00F6F8F8&
         Height          =   1125
         Left            =   75
         TabIndex        =   9
         Top             =   105
         Width           =   6795
         Begin TDBText6Ctl.TDBText txtFullname 
            Height          =   285
            Left            =   1650
            TabIndex        =   10
            Top             =   510
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   503
            Caption         =   "frmSysSignatory.frx":0000
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSysSignatory.frx":006C
            Key             =   "frmSysSignatory.frx":008A
            BackColor       =   -2147483643
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
            MaxLength       =   50
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
         Begin TDBText6Ctl.TDBText txtSigcode 
            Height          =   285
            Left            =   1650
            TabIndex        =   13
            Top             =   195
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            Caption         =   "frmSysSignatory.frx":00CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSysSignatory.frx":013A
            Key             =   "frmSysSignatory.frx":0158
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   4210752
            ReadOnly        =   -1
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
            MaxLength       =   50
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
         Begin VB.Label Label1 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   225
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Fullname"
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
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   540
            Width           =   1455
         End
      End
      Begin TrueOleDBGrid80.TDBGrid tdgSignatory 
         Height          =   4335
         Left            =   90
         TabIndex        =   12
         Top             =   1260
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Code"
         Columns(0).DataField=   "sigcode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Fullname"
         Columns(1).DataField=   "Fullname"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4260"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4180"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=260"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   2
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
         _StyleDefs(22)  =   "Splits(0).Style:id=59,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=68,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=60,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=61,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=62,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=64,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=63,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=65,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=66,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=67,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=69,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=70,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=16,.parent=59"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=60"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=61"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=63"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=86,.parent=59"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=60,.alignment=0"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=61"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=63"
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
   End
   Begin CitronSoftwarePayroll.b8ChildTitleBar titlebar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   609
      BackColor       =   12735512
      Caption         =   "Signatory"
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
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1032
      BorderColor     =   16777215
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   1
         Left            =   1185
         TabIndex        =   2
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Edit"
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
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&New"
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
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   2
         Left            =   2340
         TabIndex        =   4
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Delete"
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
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   3
         Left            =   3495
         TabIndex        =   5
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Cancel"
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
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   4
         Left            =   4650
         TabIndex        =   6
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Print"
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
         cBack           =   16185592
      End
      Begin lvButton.lvButtons_H cmdMenu 
         Height          =   420
         Index           =   5
         Left            =   5805
         TabIndex        =   7
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         Caption         =   "&Close"
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
         cBack           =   16185592
      End
   End
End
Attribute VB_Name = "frmSysSignatory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsSignatory     As ADODB.Recordset

Private Sub Form_Load()

    NetOpen rsSignatory, "", "select * from signatory order by fullname"
    
    Set tdgSignatory.DataSource = rsSignatory
    
    With rsSignatory
        If .RecordCount > 0 Then
            Lock_Button "TTFFTT", cmdMenu, 5
            txtSigcode.Text = !sigcode
            txtFullname.Text = !Fullname
        Else
            Lock_Button "TFFFFT", cmdMenu, 5
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    
    TitleBar.Move 0, 0, Me.ScaleWidth
    
    With fraButton
      .Top = TitleBar.Height
      .Left = 0
      .Width = Me.ScaleWidth
    End With
        
    With fra1
        .Top = fraButton.Top + fraButton.Height - 80
    End With

End Sub

Private Sub cmdmenu_Click(Index As Integer)
  Select Case Index
    Case 0: AddSave_Button_Clicked
    Case 1: EditUpdate_Button_Clicked
    Case 2:
    Case 3: Cancel_Clicked
    Case 4:
    Case 5: Unload Me
  End Select
End Sub

Private Sub AddSave_Button_Clicked()

  If cmdMenu(0).Caption = "&New" Then
  
    Lock_Button "TFFTFF", cmdMenu, 5
    cmdMenu(0).Caption = "&Save"
    tdgSignatory.Enabled = False
    Clear_Fields
    
  Else
    
    If Trim(txtFullname.Text) = "" Then
        MsgBox "Fullname is blank.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    txtSigcode.Text = LastCode("GetLastCodeA", "Signatory", "0000000")
    
    CitronPayroll.Execute "insert into signatory(sigcode,fullname) values " & _
                "('" & txtSigcode.Text & "','" & txtFullname.Text & "')"
    rsSignatory.Requery
    rsSignatory.MoveFirst
    rsSignatory.Find "sigcode = '" & txtSigcode.Text & "'"
    cmdmenu_Click 3
      
  End If
  
End Sub

Private Sub EditUpdate_Button_Clicked()

    Dim rsChk       As ADODB.Recordset

  If cmdMenu(1).Caption = "&Edit" Then
    
      Lock_Button "FTFTFF", cmdMenu, 5
      tdgSignatory.Enabled = False
      cmdMenu(1).Caption = "&Update"
  
  Else
  
    If Trim(txtFullname.Text) = "" Then
        MsgBox "Fullname is blank.", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    CitronPayroll.Execute "update signatory set fullname = '" & txtFullname.Text & "' where sigcode = '" & txtSigcode.Text & "'"
  
    rsSignatory.Requery
    rsSignatory.MoveFirst
    rsSignatory.Find "sigcode = '" & txtSigcode.Text & "'"
    cmdmenu_Click 3
    
  End If
  
End Sub

Private Sub Cancel_Clicked()

  If rsSignatory.RecordCount > 0 Then
    Lock_Button "TTTFTT", cmdMenu, 5
  Else
    Lock_Button "TFFFTT", cmdMenu, 5
  End If
  
  cmdMenu(0).Caption = "&New"
  cmdMenu(1).Caption = "&Edit"
  tdgSignatory.Enabled = True
  tdgSignatory_RowColChange 0, 0

End Sub

Private Sub Clear_Fields()
    txtSigcode.Text = ""
    txtFullname.Text = ""
End Sub

Private Sub tdgSignatory_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With rsSignatory
        If .RecordCount > 0 Then
            txtSigcode.Text = !sigcode
            txtFullname.Text = !Fullname
        Else
            Clear_Fields
        End If
    End With
    
End Sub




