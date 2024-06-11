VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{B168897A-CA15-457E-820F-FADB493B3E6C}#1.0#0"; "xpthing.ocx"
Begin VB.Form rptReportViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTab 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   7110
      TabIndex        =   0
      Top             =   5460
      Width           =   7110
      Begin VB.Timer Timer 
         Left            =   4350
         Top             =   180
      End
      Begin OsenXPCntrl.OsenXPButton cmdPrint 
         Height          =   465
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "&Print Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "rptReportViewer.frx":0000
         PICN            =   "rptReportViewer.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer xrpt 
      Height          =   5385
      Left            =   -15
      TabIndex        =   2
      Top             =   -45
      Width           =   10890
      lastProp        =   600
      _cx             =   19209
      _cy             =   9499
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "rptReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RS As ADODB.Recordset
Public sqlQuery As String
Public sqlWhere As String
Public sqlOrder As String

Public ReportFile As String

Dim CrxApp              As CRAXDRT.Application
Dim CrxRep              As CRAXDRT.Report
Dim crxDatabase         As CRAXDRT.Database
Dim crxDatabaseTables   As CRAXDRT.DatabaseTables
Dim crxDatabaseTable    As CRAXDRT.DatabaseTable


Public Sub Initialize()
Set CrxRep = New CRAXDRT.Report
Set CrxApp = CreateObject("crystalruntime.application")
Set CrxRep = CrxApp.OpenReport(App.Path & "\crystal\" & ReportFile)
Set crxDatabase = CrxRep.Database
Set crxDatabaseTables = crxDatabase.Tables

'For Each dbTable In CrxRep.Database.Tables
' dbTable.SetLogOnInfo SQLServerName, SQLDatabase, SQLUsername, SQLPassword
'Next dbTable

NetOpen RS, sqlQuery & sqlWhere & sqlOrder

With CrxRep
End With

CrxRep.DiscardSavedData
CrxRep.Database.SetDataSource RS

xrpt.ReportSource = CrxRep
xrpt.ViewReport

Do While xrpt.IsBusy
   DoEvents
   cmdPrint.Enabled = False
Loop

cmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
Openforms = Openforms + 1
Call Initialize
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
    xrpt.Width = Me.ScaleWidth
        If Me.xrpt.Top < Me.ScaleHeight Then
            xrpt.Height = Me.ScaleHeight - picTab.Height
        End If
    End If
End Sub

'Private Sub cmdPrint_Click()
'    If ApprovedLoans.RecordCount > 0 Then
'        CrxRep.PrinterSetup Me.hwnd
'        xrpt.PrintReport
'    End If
'End Sub


Private Sub Form_Unload(Cancel As Integer)
Openforms = Openforms - 1
Set rptReportViewer = Nothing
End Sub

Private Sub xrpt_PrintButtonClicked(UseDefault As Boolean)
    CrxRep.PrinterSetup Me.hwnd
End Sub
