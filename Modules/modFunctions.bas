Attribute VB_Name = "modFunctions"
Option Explicit

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'sentence case
Public Function cSentenceCase(sText As String) As String
    
    Dim splitText() As String
    Dim newWord As String
    Dim I As Integer
    
    'check if null---------------
    If Len(sText) < 1 Then
        cSentenceCase = ""
        Exit Function
    End If
    'end Null --------------------
    
    'convert
    sText = Trim(sText)
    
    splitText = Split(sText, " ")
    
    For I = 0 To UBound(splitText)
        If Len(Trim(splitText(I))) > 0 Then
            newWord = UCase(Left(Trim(splitText(I)), 1)) & LCase(Right(Trim(splitText(I)), Len(Trim(splitText(I))) - 1))
            cSentenceCase = cSentenceCase & " " & newWord
        End If
    Next
    
    cSentenceCase = Trim(cSentenceCase)
End Function


'Function used to format recordset
Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function

'Function used to right split user fields
Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim I As Integer
    Dim t As String
    For I = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, I, 1)
    Next I
    RightSplitUF = t
    I = 0
    t = ""
End Function

'Function used to left split user fields
Public Function LeftSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then LeftSplitUF = "": Exit Function
    Dim I As Integer
    Dim t As String
    For I = 1 To Len(srcUF)
        If Mid$(srcUF, I, 7) = "*~~~~~*" Then
            Exit For
        Else
            t = t & Mid$(srcUF, I, 1)
        End If
    Next I
    LeftSplitUF = t
    I = 0
    t = ""
End Function

'Function used to check if the record exit or not.
Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional sStr2, Optional isString As Boolean) As Boolean
    Dim RS As New Recordset
    RS.CursorLocation = adUseClient
    If isString = False Then
        RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "' " & sStr2, ConMain, adOpenStatic, adLockOptimistic
    Else
        RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "' " & sStr2, ConMain, adOpenStatic, adLockOptimistic
    End If
    If RS.RecordCount < 1 Then
        isRecordExist = False
    Else
        isRecordExist = True
    End If
    Set RS = Nothing
End Function

Public Function LastDayOfMonthAndYear(iMonth As Integer, iYear As Integer) As Integer
    Dim iX As Integer


    For iX = 31 To 1 Step -1


        If (IsDate((iMonth & "/" & iX & "/" & iYear))) Then
            Exit For
        End If
    Next
    LastDayOfMonthAndYear = iX
End Function




'-Function used to check if the record exist or not.
Public Function rec_exist_for_ae(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, ByRef sEntryField) As Boolean
Dim RS As New ADODB.Recordset
RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", ConMain, adOpenStatic, adLockReadOnly
If RS.RecordCount < 1 Then
    rec_exist_for_ae = False
Else
    MsgBox "The adding of new entry cannot be done because the PN Number already" & vbCrLf & "exist in the record.Please check and change it." & vbCrLf & vbCrLf & "Note: Duplication of entries is not allowed in this application.", vbExclamation
    sEntryField.SetFocus
    sEntryField.Text = ""
    rec_exist_for_ae = True
End If
Set RS = Nothing
End Function
'-Function used to check if the record exit or not.
Public Function rec_exist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional isNum As Boolean) As Boolean
Dim RS As New ADODB.Recordset
If isNum = True Then
    RS.Open "Select * From " & sTable & " Where " & sField & " = " & sStr, ConMain, adOpenStatic, adLockReadOnly
Else
    RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", ConMain, adOpenStatic, adLockReadOnly
End If
If RS.RecordCount < 1 Then
    rec_exist = False
Else
    rec_exist = True
End If
Set RS = Nothing
End Function

'-Function used to check if the field is empty or not.
Public Function is_empty(ByRef sText As Variant) As Boolean
If sText.Text = "" Then
    is_empty = True
    MsgBox "All fields must be filled up. Please check it!", vbExclamation
    sText.SetFocus
Else
    is_empty = False
End If
End Function

Public Function rec_found(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal sFindText As String) As Boolean
'-Move the recordset to the first record
sRS.Requery '-Use this instead of movefirst so that new record added can be used immediately
'Search the record
sRS.Find sField & " = '" & sFindText & "'"
'-Verify if the search string was found or not
If sRS.EOF Then
    rec_found = False
Else
    rec_found = True
End If
End Function

Public Function getCount(ByVal srctable As String) As Long
Dim RS As New ADODB.Recordset
With RS
    .Open "SELECT * FROM " & srctable & "", ConMain, adOpenStatic, adLockReadOnly
    getCount = .RecordCount
End With
Set RS = Nothing
End Function

Public Function dotsCounter(ByVal sStr As String) As Long
If sStr = "" Then Exit Function
Dim C As Long
For C = 1 To Len(sStr)
   If Mid$(sStr, C, 1) = "." Then dotsCounter = dotsCounter + 1
Next C
C = 0
End Function

Public Function getMax(ByVal sTable As String, ByVal sField As String) As Long
On Error GoTo err
Dim RS As New ADODB.Recordset
RS.Open "SELECT Max(" & sTable & "." & sField & ") AS [Number] From " & sTable & " ORDER BY Max(" & sTable & "." & sField & ") DESC", ConMain, adOpenStatic, adLockOptimistic
getMax = RS.Fields(0)

sTable = ""
sField = ""
Set RS = Nothing
Exit Function
err:
    'Error when incounter a null value
    'If err.Number = 94 Then get_num = 1: Resume Next
End Function

Public Function getRecCount(ByVal srcSQL As String) As Long
Dim RS As New ADODB.Recordset
With RS
    .Open srcSQL, ConMain, adOpenStatic, adLockReadOnly
    getRecCount = .RecordCount
End With
Set RS = Nothing
End Function

Function Amt2Words(nInAmount As Double) As String
    Dim sInWords As String, snum As String, nCent As Double, sThree As String
    Dim sNum1 As String, nCtr As Integer, sWord As String, lcont As Boolean
    
    Dim aTens(9) As String, aOnes(9) As String, aCValue(9) As String
    Dim nLen As Integer, X As Integer, nSingle As Integer
    
    
    aOnes(1) = "One"
    aOnes(2) = "Two"
    aOnes(3) = "Three"
    aOnes(4) = "Four"
    aOnes(5) = "Five"
    aOnes(6) = "Six"
    aOnes(7) = "Seven"
    aOnes(8) = "Eight"
    aOnes(9) = "Nine"

    aTens(1) = "Ten"
    aTens(2) = "Twenty"
    aTens(3) = "Thirty"
    aTens(4) = "Fourty"
    aTens(5) = "Fifty"
    aTens(6) = "Sixty"
    aTens(7) = "Seventy"
    aTens(8) = "Eigthy"
    aTens(9) = "Ninety"

    aCValue(1) = "Eleven"
    aCValue(2) = "Twelve"
    aCValue(3) = "Thirteen"
    aCValue(4) = "Fourteen"
    aCValue(5) = "Fifteen"
    aCValue(6) = "Sixteen"
    aCValue(7) = "Seventeen"
    aCValue(8) = "Eighteen"
    aCValue(9) = "Nineteen"
    
    nInAmount = Abs(nInAmount)
    snum = Trim(str(Int(nInAmount)))
    nCent = 0
    If Val(snum) > 0 Then
        nCent = nInAmount - Val(snum)
    Else
        nCent = nInAmount
    End If

    nCent = nCent * 100
    nLen = Len(snum)
    If nLen < 12 Then
        sNum1 = Stuff(snum, 1, "0", 12 - Len(snum))
    Else
        sNum1 = snum
    End If
    sInWords = ""
    
    nCtr = 1
    Do While True
        sThree = Mid(sNum1, nCtr, 3)
        sWord = ""
        For X = 1 To 3
            nSingle = Val(Mid(sThree, X, 1))
            lcont = True
            If nSingle > 0 Then
                If X = 1 Then
                    sWord = sWord + aOnes(nSingle) + " Hundred "
                End If
                If X = 2 Then
                    If nSingle = 1 And Val(Mid(sThree, 3, 1)) > 0 Then
                        sWord = sWord + " " + aCValue(Val(Mid(sThree, 3, 1)))
                        lcont = False
                    Else
                        If nSingle > 0 Then
                            sWord = sWord + " " + aTens(nSingle)
                        End If
                    End If
                End If
            
                If Not lcont Then
                    Exit For
                End If
                If X = 3 Then
                    sWord = sWord + " " + aOnes(nSingle)
                End If
            End If
        Next X
    
        sInWords = sInWords + " " + sWord
        If nCtr = 1 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Billion"
        End If
    
        If nCtr = 4 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Million"
        End If
    
        If nCtr = 7 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords & " " & "Thousand"
        End If
    
        nCtr = nCtr + 3
        If nCtr > 13 Then
            Exit Do
        End If
    
    Loop
    
    'I use Peso coz its our currency name in the Philippines
    'Just change it whatever currency word you have...
    
    If Val(snum) > 1 Then
        sInWords = sInWords & " " & "Pesos"
    End If
    
    If Val(snum) = 1 Then
        sInWords = sInWords + " " + "Peso"
    End If
    
    nCent = Format(nCent, "0.00")
    
    If nCent > 0 And Val(snum) > 1 Then
        sInWords = sInWords + " " + "and" + " " + Trim(str(nCent)) + "/100"
    End If

    If nCent > 0 And Val(snum) = 0 Then
        sInWords = sInWords + " " + Trim(str(nCent)) + "/100"
    End If
    
    sInWords = sInWords
    Amt2Words = Trim(sInWords)
End Function

Function Percent2Words(nInAmount As Double) As String
    Dim sInWords As String, snum As String, nCent As Double, sThree As String
    Dim sNum1 As String, nCtr As Integer, sWord As String, lcont As Boolean
    
    Dim aTens(9) As String, aOnes(9) As String, aCValue(9) As String
    Dim nLen As Integer, X As Integer, nSingle As Integer
    
    
    aOnes(1) = "One"
    aOnes(2) = "Two"
    aOnes(3) = "Three"
    aOnes(4) = "Four"
    aOnes(5) = "Five"
    aOnes(6) = "Six"
    aOnes(7) = "Seven"
    aOnes(8) = "Eight"
    aOnes(9) = "Nine"

    aTens(1) = "Ten"
    aTens(2) = "Twenty"
    aTens(3) = "Thirty"
    aTens(4) = "Fourty"
    aTens(5) = "Fifty"
    aTens(6) = "Sixty"
    aTens(7) = "Seventy"
    aTens(8) = "Eigthy"
    aTens(9) = "Ninety"

    aCValue(1) = "Eleven"
    aCValue(2) = "Twelve"
    aCValue(3) = "Thirteen"
    aCValue(4) = "Fourteen"
    aCValue(5) = "Fifteen"
    aCValue(6) = "Sixteen"
    aCValue(7) = "Seventeen"
    aCValue(8) = "Eighteen"
    aCValue(9) = "Nineteen"
    
    nInAmount = Abs(nInAmount)
    snum = Trim(str(Int(nInAmount)))
    nCent = 0
    If Val(snum) > 0 Then
        nCent = nInAmount - Val(snum)
    Else
        nCent = nInAmount
    End If

    nCent = nCent * 100
    nLen = Len(snum)
    If nLen < 12 Then
        sNum1 = Stuff(snum, 1, "0", 12 - Len(snum))
    Else
        sNum1 = snum
    End If
    sInWords = ""
    
    nCtr = 1
    Do While True
        sThree = Mid(sNum1, nCtr, 3)
        sWord = ""
        For X = 1 To 3
            nSingle = Val(Mid(sThree, X, 1))
            lcont = True
            If nSingle > 0 Then
                If X = 1 Then
                    sWord = sWord + aOnes(nSingle) + " Hundred "
                End If
                If X = 2 Then
                    If nSingle = 1 And Val(Mid(sThree, 3, 1)) > 0 Then
                        sWord = sWord + " " + aCValue(Val(Mid(sThree, 3, 1)))
                        lcont = False
                    Else
                        If nSingle > 0 Then
                            sWord = sWord + " " + aTens(nSingle)
                        End If
                    End If
                End If
            
                If Not lcont Then
                    Exit For
                End If
                If X = 3 Then
                    sWord = sWord + " " + aOnes(nSingle)
                End If
            End If
        Next X
    
        sInWords = sInWords + " " + sWord
        If nCtr = 1 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Billion"
        End If
    
        If nCtr = 4 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Million"
        End If
    
        If nCtr = 7 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords & " " & "Thousand"
        End If
    
        nCtr = nCtr + 3
        If nCtr > 13 Then
            Exit Do
        End If
    
    Loop
    
    'I use Peso coz its our currency name in the Philippines
    'Just change it whatever currency word you have...
    
    If Val(snum) > 1 Then
        sInWords = sInWords & " " & "Percent"
    End If
    
    If Val(snum) = 1 Then
        sInWords = sInWords + " " + "Percent"
    End If
    
    nCent = Format(nCent, "0.00")
    
    If nCent > 0 And Val(snum) > 1 Then
        sInWords = sInWords + " " + "and" + " " + Trim(str(nCent)) + "/100"
    End If

    If nCent > 0 And Val(snum) = 0 Then
        sInWords = sInWords + " " + Trim(str(nCent)) + "/100"
    End If
    
    sInWords = sInWords
    Percent2Words = Trim(sInWords)
    
End Function


'Parameters: 1. sStr : String to be stuff
'            2. cPos : Position where it is inserted
'                      1 : Left
'                      2 : Right
'            3. cStuff: Character to be stuff
'            4. nNo   : how many times

Function Stuff(sStr, cPos As Byte, cStuff As String, nNo As Byte) As String

    Dim sString As String, X As Byte
    sString = ""
    For X = 1 To nNo
        sString = sString & cStuff
    Next X
    If cPos = 1 Then
        sString = sString & sStr
    End If
    
    If cPos = 2 Then
        sString = sStr & sString
    End If
    
    Stuff = sString
    
End Function

Public Function LCaseKeyPress(ByRef KeyAscii As Integer) As Integer
    ' Useful in the KeyPress event to convert entry to LCase()
    LCaseKeyPress = Asc(LCase(Chr(KeyAscii)))
End Function


Public Function UCaseKeyPress(ByRef KeyAscii As Integer) As Integer
    UCaseKeyPress = Asc(UCase(Chr(KeyAscii)))
End Function

'Function Round(RoundMe, RoundTo)
'    Round = Int((RoundMe * 10 ^ RoundTo) + 0.5) / 10 ^ RoundTo
'End Function


Public Function WHOLENUM(n As Double)
  Dim I     As Integer
  Dim mSTR  As String
  I = 1
  For I = 1 To Len(CStr(n))
      If Mid(CStr(n), I, 1) <> "." Then
        mSTR = mSTR & Mid(CStr(n), I, 1)
      Else
        Exit For
      End If
  Next
  WHOLENUM = CInt(mSTR)
End Function

Public Function CAPS(objectz As Object)
    objectz.Text = UCase(objectz.Text)
    objectz.SelStart = Len(objectz.Text)
End Function

Public Function SCAPS(objectz As Object)
    objectz.Text = LCase(objectz.Text)
    objectz.SelStart = Len(objectz.Text)
End Function

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srctable As String, ByVal Fieldz As String, Optional srccondition As String, Optional isformatted As Boolean) As String
    If srccondition <> "" Then srccondition = " " & srccondition
    Dim RS As New Recordset
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT COUNT('" & Fieldz & "') as TCount FROM " & srctable & srccondition, ConMain, adOpenStatic, adLockReadOnly
    If isformatted = True Then
        getRecordCount = Format$(RS![TCount], "#,##0")
    Else
        getRecordCount = RS![TCount]
    End If
    Set RS = Nothing
End Function



Function CSQ(ByVal str) As String
     If IsNull(str) Then str = ""
     CSQ = Replace(str, "'", "''")
End Function

Public Sub SafeSetFocus(ByVal ctlFocus As Control)
On Error Resume Next
If ctlFocus.Visible And ctlFocus.Enabled Then
   ctlFocus.SetFocus
End If
End Sub

Public Sub BindDropDown(ByVal dropdown As TDBDropDown, ByVal sqlString As String, Optional ListFieldString As String)
Dim RS As New ADODB.Recordset
NetOpen RS, sqlString
If RS.RecordCount > 0 Then
    Set dropdown.DataSource = RS
    If ListFieldString <> "" Then
        dropdown.ListField = ListFieldString
    End If
Else
    Set dropdown.DataSource = RS
    If ListFieldString <> "" Then
        dropdown.ListField = ListFieldString
    End If
End If
End Sub

Public Function DuplicateCheck(ByVal tdbgrid As tdbgrid, ByVal Value As String, ByVal Column As Long) As Boolean
On Error Resume Next
Dim I As Long
Dim count As Long
Dim dup As Long
count = tdbgrid.ApproxCount - 1
dup = 0
If tdbgrid.Row > 0 Then
For I = 0 To count
    tdbgrid.Row = I
    If Trim(Value) = Trim(tdbgrid.Columns(Column).Value) Then
        dup = dup + 1
    End If
Next I
End If
If dup > 1 Then
    DuplicateCheck = True
Else
    DuplicateCheck = False
End If
End Function

Public Function GetFileExt(mfile As String) As String

    Dim I           As Integer
    
    I = Len(mfile)
    Do While I > 0
        If Mid(mfile, I, 1) <> "." Then
            GetFileExt = Right(mfile, Len(mfile) - (I - 1))
            I = I - 1
        Else
            Exit Do
        End If
    Loop

End Function
