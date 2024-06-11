Attribute VB_Name = "modGeneralVariable"
Option Explicit

'Software Info
Public ModuleVersion         As String

'Database connections
Public ConMain                As New ADODB.Connection
Public ConAdvPayroll          As New ADODB.Connection
Public Cnstr                  As New ADODB.Connection

'userinfo
Public GlobalUserID           As Integer
Public UserName               As String
Public UserType               As String
Public GlobalUserGroupID      As String

'CompanyInfo
Public CompanyName      As String
Public Address          As String
Public Telephone        As String
Public EmailAdd         As String

'File Paths
Public mEmpPicPath      As String

'For Reports
Public mReport          As New CRAXDRT.Report


Public mReports         As CRAXDRT.Report
Public mReportPath      As CRAXDRT.Application

Public SQLServerName    As String
Public SQLDatabase      As String
Public SQLUsername      As String
Public SQLPassword      As String
Public SQLPort          As String

Public Openforms        As Long

Public mLogOn           As Boolean
