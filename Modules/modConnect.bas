Attribute VB_Name = "modConnect"
Option Explicit

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                                                                            ByVal lpFileName As String, _
                                                                                            ByVal nSize As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                          ByVal lpszShortPath As String, _
                                                                                          ByVal lBuffer As Long) As Long
                                                                                            
Public Sub Main()

'    Dim FileName1       As String
'    Dim txt             As String
'    Dim Phrase          As String
'    Dim Char1           As String
'    Dim mEncrypted      As String
'    Dim mSourcePath     As String
'    Dim mString         As String
'    Dim mOldApp         As String
'
'    Dim Position        As Integer
'    Dim ctr             As Integer
'
'    Dim Asc1            As Long
'
'    Dim mToUpdate       As Boolean
'
'    mToUpdate = False
'
'    FileName1 = App.Path & "\ExtData\upd8.dat"
'
'    If DirExists(FileName1) = True Then
'
'        mEncrypted = ""
'
'        Open FileName1 For Input As #1
'            Do Until EOF(1)
'                Input #1, txt
'                mEncrypted = mEncrypted & txt
'            Loop
'        Close #1
'
'        Phrase = mEncrypted
'
'        mEncrypted = ""
'
'        ctr = 0
'
'        For Position = Len(Phrase) To 1 Step -1
'
'            Char1 = Mid$(Phrase, Position, 1)
'            Asc1 = Asc(Char1)
'            Asc1 = (((Asc1 * Asc1) / 2) / 2)
'            Asc1 = Sqr(Asc1)
'            Char1 = Chr$(Asc1)
'
'            If Char1 = "," Then
'                Char1 = ""
'                ctr = ctr + 1
'                If ctr = 1 Then
'                    mSourcePath = mEncrypted
'                    mEncrypted = ""
'                ElseIf ctr = 2 Then
'                    If mEncrypted = 1 Then
'                        mToUpdate = True
'                    End If
'                    mEncrypted = ""
'                ElseIf ctr = 3 Then
'                    mOldApp = mEncrypted
'                    mEncrypted = ""
'                End If
'            End If
'
'            mEncrypted = mEncrypted & Char1
'
'        Next
'
'    End If
'
'    If DirExists(mSourcePath) = True And Trim(mSourcePath) <> "" Then
'        If mSourcePath <> App.Path & "\" & App.EXEName & ".exe" Then
'          If Update(mSourcePath) = False Then
'            frmLogin.Show
'          End If
'        Else
            'frmLogin.Show
'        End If
'    Else
'        frmLogin.Show
'    End If

End Sub

Public Function Update(m_RemoteEXE As String) As Boolean
 
            Dim b()                     As Byte
            
            Dim strLocalNum()           As String
            Dim strRemoteNum()          As String
            Dim strLocalVer             As String
            Dim strRemoteVer            As String
            Dim strFileName             As String
            Dim strPath                 As String
            Dim strFile                 As String
            Dim Phrase                  As String
            Dim mString                 As String
            Dim Char1                   As String
            Dim FileName1               As String
            
            Dim blnUpdate               As Boolean
            
            Dim fsoFile                 As File
            
            Dim FSO                     As FileSystemObject
            
            Dim Position                As Integer
            Dim ctr                     As Integer
            
            Dim Asc1                    As Long
            Dim I                       As Long
 
          'Get full path of parent exe (i.e. the local file)
 
          strFileName = Space$(255)
 
          Call GetModuleFileName(GetModuleHandle(vbNullString), strFileName, Len(strFileName))
 
          strFileName = Split(strFileName, vbNullChar)(0)
 
          strPath = returnPathOfFile(strFileName)

          strFile = returnNameOfFile(strFileName)
          'Get the version number of the local file

          Set FSO = New FileSystemObject

          Set fsoFile = FSO.GetFile(strFileName)

          strLocalVer = FSO.GetFileVersion(strFileName)
          'Get the version number of the remote file

          'strFileName = Right(m_RemoteEXE, Len(m_RemoteEXE) - 7)
            strFileName = m_RemoteEXE
            
          Set FSO = New FileSystemObject
 
          Set fsoFile = FSO.GetFile(strFileName)
 
          strRemoteVer = FSO.GetFileVersion(strFileName)
 
          'Compare version numbers
 
          If strRemoteVer = strLocalVer Then
 
              blnUpdate = False
 
          Else
 
              strRemoteNum() = Split(strRemoteVer, ".")
 
              strLocalNum() = Split(strLocalVer, ".")
 
              'Compare major, then minor, then revision
 
              For I = 0 To UBound(strRemoteNum)

                  If CInt(strRemoteNum(I)) > CInt(strLocalNum(I)) Then

                        If MsgBox("A more recent version of this program exists. Would you like to update it now?", vbYesNo Or vbQuestion) = vbYes Then

                              blnUpdate = True

                        Else

                              blnUpdate = False

                        End If

                        Exit For

                  ElseIf CInt(strRemoteNum(I)) < CInt(strLocalNum(I)) Then

                      blnUpdate = False

                      Exit For

                  Else

                      'ie values are the same

                      blnUpdate = False

                  End If

              Next

          End If
 
          'If blnUpdate = True, then download the latest program exe from the remote site
 
            If blnUpdate Then
            
                Phrase = m_RemoteEXE & "," & 1 & "," & App.EXEName & ".exe" & ","
                mString = ""
                For Position = Len(Phrase) To 1 Step -1
                    Char1 = Mid$(Phrase, Position, 1)
                    Asc1 = Asc(Char1)
                    Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
                    Char1 = Chr$(Asc1)
                    mString = mString & Char1
                Next
                
                FileName1 = App.Path & "\Extdata\upd8.dat"
                
                Open FileName1 For Output As #1
                    Print #1, mString
                Close #1
                
            
                Shell (App.Path & "\Updater.exe")
            
                DestroyAllObjects
                
                End
            
            Else
'
                Update = False
            
            End If
 
End Function

Private Function GetShortPath(strFileName As String) As String


    Dim lngRes As Long, strPath As String
    
    'Create a buffer
    
    strPath = String$(165, 0)
    
    'retrieve the short pathname
    
    lngRes = GetShortPathName(strFileName, strPath, 164)
    
    'remove all unnecessary chr$(0)'s
    
    GetShortPath = Left$(strPath, lngRes)

End Function

' Return the path of a given full path to a file
Private Function returnPathOfFile(ByVal strFile As String) As String

    returnPathOfFile = Left(strFile, InStrRev(strFile, "\"))

End Function

' Return the filename of a given full path to a file
Private Function returnNameOfFile(ByVal strFile As String) As String

    returnNameOfFile = Mid(strFile, InStrRev(strFile, "\") + 1)

End Function

