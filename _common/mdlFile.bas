Attribute VB_Name = "mdlFile"
Option Explicit

Public bExitFlag As Boolean

Private Const GREP_ARRAY_LEN As Long = 9999

''
' Legge un file
Public Function ReadFile(ByVal NomeFile As String) As String
On Error GoTo err1
Dim numfile As Long

    numfile = FreeFile
    Open NomeFile For Input As #numfile
    Line Input #numfile, ReadFile
    Close #numfile
    Exit Function
    
err1:
    ReadFile = ""
End Function

''
'
Public Sub FileFinder(ByRef CallbackForm As Form, ByRef sBaseDir As String, ByVal sBaseFile As String, ByVal bSearchSubdir As Boolean)
On Error Resume Next
Dim h As Long
Dim fd As WIN32_FIND_DATA
Dim sFileName As String

    If Right(sBaseDir, 1) <> "\" Then
        sBaseDir = sBaseDir & "\"
    End If
    
    '------------- ricerca pattern
    h = FindFirstFile(sBaseDir & sBaseFile, fd)
    Do
        If h <= 0 Then Exit Do
        sFileName = Mid(fd.cFileName, 1, InStr(fd.cFileName, Chr(0)) - 1)
        If sFileName = "" Then Exit Do
        
        If sFileName <> "." And sFileName <> ".." Then
            CallbackForm.ElaboraFile sBaseDir & sFileName, ((fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
        End If
        
        DoEvents
        If bExitFlag Then
            FindClose h
            Exit Sub
        End If
    Loop While FindNextFile(h, fd)
    FindClose h
    
    '------------- ricerca subdirectory
    If bSearchSubdir Then
        h = FindFirstFile(sBaseDir & "*.", fd)
        While FindNextFile(h, fd)
            'ricerca delle subdir
            sFileName = Mid(fd.cFileName, 1, InStr(fd.cFileName, Chr(0)) - 1)
            If sFileName <> "." And sFileName <> ".." Then
                If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
                    FileFinder CallbackForm, sBaseDir & sFileName, sBaseFile, True
                End If
            End If
            
            DoEvents
            If bExitFlag Then
                FindClose h
                Exit Sub
            End If
        Wend
        FindClose h
    End If
End Sub

''
'
Public Function GrepFile(ByVal sFileName As String, ByVal bMatchCase As Boolean, ByRef aFindArray() As Byte, ByRef aFindArrayMC() As Byte) As Boolean
On Error GoTo err1
Dim iFileNum As Integer
Dim tmp As Byte
Dim iCountFind As Integer
Dim bFindByte As Boolean

    iFileNum = FreeFile
    iCountFind = 0
    Open sFileName For Binary As #iFileNum
    
    While Not EOF(iFileNum)
        Get #iFileNum, , tmp
        
        If bMatchCase Then
            bFindByte = aFindArray(iCountFind) = tmp
        Else
            bFindByte = aFindArray(iCountFind) = tmp Or aFindArrayMC(iCountFind) = tmp
        End If
        
        If bFindByte Then
            If iCountFind = UBound(aFindArray) Then
                GrepFile = True
                Close #iFileNum
                Exit Function
            End If
            iCountFind = iCountFind + 1
        Else
            iCountFind = 0
        End If
        bFindByte = False
        
        If bExitFlag Then
            Close #iFileNum
            Exit Function
        End If
        DoEvents
    Wend
    
    Close #iFileNum
    Exit Function
    
err1:
    MsgBox Err.Description & " (" & sFileName & ")"
End Function

''
' NB: gli array devono partire tutti da 0
Public Function GrepFile2(ByVal sFileName As String, ByVal bMatchCase As Boolean, ByVal iArrayLen As Integer, ByRef aFindArray() As Byte, ByRef aFindArrayMC() As Byte) As Boolean
On Error GoTo err1
Dim iFileNum As Integer
Dim aRead(0 To GREP_ARRAY_LEN) As Byte 'lettura a blocchi
Dim iCountFind As Integer
Dim bFindByte As Boolean

    iFileNum = FreeFile
    iCountFind = 0
    Open sFileName For Binary As #iFileNum
    
    While Not EOF(iFileNum)
        Get #iFileNum, , aRead
        
        If bMatchCase Then
            iCountFind = isInArray(aRead, GREP_ARRAY_LEN, aFindArray, aFindArrayMC, iCountFind, iArrayLen, False)
        Else
            iCountFind = isInArray(aRead, GREP_ARRAY_LEN, aFindArray, aFindArrayMC, iCountFind, iArrayLen, True)
        End If
        
        If iCountFind = iArrayLen Then
            GrepFile2 = True
            Close #iFileNum
            Exit Function
        End If
            
        If bExitFlag Then
            Close #iFileNum
            Exit Function
        End If
        DoEvents
    Wend
    
    Close #iFileNum
    Exit Function
    
err1:
    MsgBox Err.Description & " (" & sFileName & ")"
End Function

''
' NB: gli array devono partire tutti da 0
Public Function GrepFile3(ByVal sFileName As String, ByVal bMatchCase As Boolean, ByVal iArrayLen As Integer, ByRef aFindArray() As Byte, ByRef aFindArrayMC() As Byte) As Boolean
On Error GoTo err1
Dim iFileNum As Integer
Dim aRead(0 To GREP_ARRAY_LEN) As Byte 'lettura a blocchi
Dim iCountFind As Long
Dim bFindByte As Boolean
Dim iTest As Integer

    iFileNum = FreeFile
    iCountFind = 0
    Open sFileName For Binary As #iFileNum
    
    While Not EOF(iFileNum)
        Get #iFileNum, , aRead
        
        If bMatchCase Then
            iCountFind = isInArray3(aRead, GREP_ARRAY_LEN, aFindArray, aFindArrayMC, iCountFind, iArrayLen, False)
        Else
            iCountFind = isInArray3(aRead, GREP_ARRAY_LEN, aFindArray, aFindArrayMC, iCountFind, iArrayLen, True)
        End If
        
        If iCountFind = iArrayLen Then
            GrepFile3 = True
            Close #iFileNum
            Exit Function
        End If
            
        If bExitFlag Then
            Close #iFileNum
            Exit Function
        End If
        DoEvents
    Wend
    
    Close #iFileNum
    Exit Function
    
err1:
    MsgBox Err.Description & " (" & sFileName & ")"
End Function

''
' Crea un file vuoto
Public Sub CreateFile(ByVal sFileName As String)
On Error GoTo err1
Dim lNumFile As Long

    lNumFile = FreeFile
    Open sFileName For Append As #lNumFile
    Print #lNumFile, ""
    Close #lNumFile
    
err1:
End Sub

'------------------------------------
'------------------------------------ PRIVATE
'------------------------------------

''
' NB: gli array devono partire tutti da 0
Private Function isInArray(ByRef aArr() As Byte, ByVal iAEnd As Integer, _
                          ByRef aSearch1() As Byte, ByRef aSearch2() As Byte, _
                          ByVal iSStart As Integer, ByVal iSLen As Integer, _
                          ByVal bUseSearch2 As Boolean) As Integer
On Error GoTo err1
Dim i As Integer
Dim j As Integer
Dim bFound As Boolean

    If bUseSearch2 Then 'usa anche aSearch2
        i = iSStart
        For j = 0 To iAEnd
            If aSearch1(i) = aArr(j) Or aSearch2(i) = aArr(j) Then
                bFound = True
                i = i + 1
                If i = iSLen Then
                    'parola trovata
                    isInArray = iSLen
                    Exit Function
                End If
            Else
                bFound = False
                i = 0 'non iSStart!!!
            End If
        Next j
    Else
        i = iSStart
        For j = 0 To iAEnd
            If aSearch1(i) = aArr(j) Then
                bFound = True
                i = i + 1
                If i = iSLen Then
                    'parola trovata
                    isInArray = iSLen
                    Exit Function
                End If
            Else
                bFound = False
                i = 0 'non iSStart!!!
            End If
        Next j
    End If
    
    If bFound Then
        isInArray = i
    Else
        isInArray = 0
    End If
    Exit Function
    
err1:
    isInArray = 0
End Function

''
' NB: gli array devono partire tutti da 0
Private Function isInArray3(ByRef aArr() As Byte, ByVal iAEnd As Long, _
                          ByRef aSearch1() As Byte, ByRef aSearch2() As Byte, _
                          ByVal iSStart As Long, ByVal iSLen As Long, _
                          ByVal bUseSearch2 As Boolean) As Long
On Error GoTo err1
Dim i As Long
Dim j As Long
Dim bFound As Boolean
Dim t1, t2

    If bUseSearch2 Then 'usa anche aSearch2
        i = iSStart
        For j = 0 To iAEnd
            If aSearch1(i) = aArr(j) Or aSearch2(i) = aArr(j) Then
                bFound = True
                i = i + 1
                If i = iSLen Then 'parola trovata
                    isInArray3 = iSLen
                    Exit Function
                End If
            Else
                bFound = False
                i = 0 'non iSStart!!!
            End If
        Next j
    Else
        i = iSStart
        For j = 0 To iAEnd
            If aSearch1(i) = aArr(j) Then
                bFound = True
                i = i + 1
                If i = iSLen Then
                    'parola trovata
                    isInArray3 = iSLen
                    Exit Function
                End If
            Else
                bFound = False
                i = 0 'non iSStart!!!
            End If
        Next j
    End If
    
    If bFound Then
        isInArray3 = i
    Else
        isInArray3 = 0
    End If
    Exit Function
    
err1:
    isInArray3 = 0
    Debug.Assert False
End Function
