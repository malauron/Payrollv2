Attribute VB_Name = "modNextCode"
Option Explicit

Public Function LastCode(ByVal TableName As String) As Integer
    
    Dim rs As New ADODB.Recordset
    Dim NewNum As String
    Dim I As Integer

    NetOpen rs, "select lastcodeused from lastcodeseries where module = '" & TableName & "' for update"

    If rs.RecordCount = 0 Then
        LastCode = 1
        ConMain.Execute "insert into lastcodeseries values ('" & TableName & "', 2)"
    Else
        LastCode = rs!LastCodeUsed
        ConMain.Execute "update lastcodeseries set lastcodeused = " & CInt(LastCode) + 1 & " where module = '" & TableName & "'"
    End If
        
End Function

Public Function LastCodeUsed(ByVal ColName As String, ByVal mPerCode As Integer) As Integer
    
    Dim rs        As New ADODB.Recordset
    Dim NewNum    As String
    Dim I         As Integer

    NetOpen rs, "select " & ColName & " lastcode from payrollperiod where percode = " & mPerCode & " for update"

    LastCodeUsed = rs!LastCode
    
    ConMain.Execute "update payrollperiod set " & ColName & " = " & CInt(LastCodeUsed) + 1 & " where percode = '" & mPerCode & "'"
    
End Function

Public Function LastLoanCodeUsed(ByVal mLoanCode As Integer) As Integer
    
    Dim rs As New ADODB.Recordset
    Dim NewNum As String
    Dim I As Integer

    NetOpen rs, "select lastloandedcode from loans where loancode = " & mLoanCode & " for update"

    LastLoanCodeUsed = rs!LastloandedCode
    ConMain.Execute "update loans set lastloandedcode = " & CInt(LastLoanCodeUsed) + 1 & " where loancode = '" & mLoanCode & "'"
    
End Function

Public Function LastReceivableCode(ByVal mCriteria As String) As Double
    
    Dim rs As New ADODB.Recordset
    Dim NewNum As String
    Dim I As Integer

    NetOpen rs, "select lastcodeused from receivablessequencecode where bases = '" & mCriteria & "' for update"

    If rs.RecordCount = 0 Then
        MsgBox CStr(Format(Weekday(CDate(mCriteria), "00")) & Format(CDate(mCriteria), "YY") & "0000001")
        LastReceivableCode = CStr(Format(CDate(mCriteria), "YY") & "000001")
        ConMain.Execute "insert into receivablessequencecode values ('" & mCriteria & "', 2)"
    Else
        LastReceivableCode = CStr(Format(CDate(mCriteria), "YY") & Format(rs!LastCodeUsed, "0000000"))
        ConMain.Execute "update receivablessequencecode set lastcodeused = " & CInt(rs!LastCodeUsed) + 1 & " where trnxdate = '" & mCriteria & "'"
    End If
        
End Function

