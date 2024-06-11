Attribute VB_Name = "modConnectDB"
'Option Explicit
'
'Public Function OpenConn(srvIP As String, dbNAme As String, dbUser As String, dbPass As String, dbPORT As String)
'
'    On Error GoTo errH
'
'    If ConMain.State <> 0 Then ConMain.Close 'Check if currently connected if yes, disconnect.
'
'    ConMain.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & srvIP & ";DATABASE=" & dbNAme & ";" _
'                                 & "UID=" & dbUser & ";PWD=" & dbPass & "; PORT=" & dbPORT & "; OPTION=3"
'    ConMain.Open  'open the connection
'
'
'    'Temporary Connections
'
''-----------------------------------------------------------
'    ConAdvPayroll.Open "Provider=Advantage OLE DB Provider;" & _
'           "Data source=D:\bliss\Shagreen\Payroll\ExtData;" & _
'           "ServerType=ADS_LOCAL_SERVER;" & _
'           "TableType=ADS_CDX"
'
'      Cnstr.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=p@ssw0rd;Initial Catalog=fairbankpayroll;Data Source=virus"
''------------------------------------------------------------
'
'    Exit Function
'errH:
'    MsgBox err.Description, vbCritical, "ERROR!"
'
'End Function
'
'
'


