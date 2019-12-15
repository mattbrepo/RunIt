Attribute VB_Name = "mdlSystem"
Option Explicit

''
' Preleva una stringa da un file INI
Public Function GetStringProp(ByVal prop As String, ByVal propDefault As String, ByVal FileName As String, Optional sectionName As String = "MAIN") As String
Dim size, rtn As Long
Dim buffer As String

    size = 128
    buffer = String(size, " ")
    rtn = GetPrivateProfileString(sectionName, prop, propDefault, buffer, size, FileName)
    If (rtn < 1) Then
        GetStringProp = propDefault
    Else
        GetStringProp = Left(buffer, rtn)
    End If
End Function

''
' Scrive una stringa da un file INI
Public Function SetStringProp(ByVal prop As String, ByVal value As String, ByVal FileName As String) As Boolean
Dim rtn As Long

    rtn = WritePrivateProfileString("MAIN", prop, value, FileName)

    If rtn < 1 Then
        SetStringProp = False
    Else
        SetStringProp = True
    End If
End Function

''
' Ritorna il path completo del browser di default
Public Function GetDefaultBrowser() As String
Dim sValue As String
    
    If GetKeyValue(HKEY_CLASSES_ROOT, "htmlfile\shell\open\command", "", sValue) Then
        GetDefaultBrowser = sValue
    End If
End Function

''
' Ritorna il path completo del mailer di default
Public Function GetDefaultMailer() As String
Dim sValue As String
    
    If GetKeyValue(HKEY_CLASSES_ROOT, "mailto\shell\open\command", "", sValue) Then
        GetDefaultMailer = sValue
    End If
End Function

''
' Ritorna il path completo del programma associato ad una data estensione
'
' @param nome dell'estensione
Public Function GetFileAssociation(ByVal sExtName As String, ByVal bWinXpSpecial As Boolean) As String
Dim sProgName As String

    If bWinXpSpecial Then
        '------- PRIMA VIA (MTB [23/01/2008]: utile per WinXP su immagini)
        If Not GetFileAssociation0(sExtName, sProgName) Then
            '------- SECONDA VIA
            If Not GetFileAssociation1(sExtName, sProgName) Then
                '------- TERZA VIA
                If Not GetFileAssociation2(sExtName, sProgName) Then
                    'MTB [16/05/2007]: qui solo per debug
                    'MsgBox sProgName
                    Exit Function
                End If
            End If
        End If
    Else
        '------- PRIMA VIA
        If Not GetFileAssociation1(sExtName, sProgName) Then
            '------- SECONDA VIA
            If Not GetFileAssociation2(sExtName, sProgName) Then
                'MTB [16/05/2007]: qui solo per debug
                'MsgBox sProgName
                Exit Function
            End If
        End If
    End If
    
    'pulizia del nome programma
    sProgName = Replace(sProgName, """", "")
    sProgName = Trim(Replace(sProgName, "%1", ""))

    '------- RECUPERO DEL PATH (SE NECESSARIO)
    If InStr(sProgName, "/") > 0 Then
        'il path e' gia' incluso
        GetFileAssociation = sProgName
    Else
        'si recupera il path dal registry
        GetFileAssociation = GetProgramPath(sProgName)
    End If
End Function

Public Function GetSpecialFolder(CSIDL As Long) As String
On Error GoTo Err_GetFolder
Dim idlstr As Long
Dim sPath As String
Dim IDL As ITEMIDLIST

   ' Fill the idl structure with the specified folder item.
   idlstr = SHGetSpecialFolderLocation(0, CSIDL, IDL)

    If idlstr = 0 Then
        ' Get the path from the idl list, and return
        ' the folder with a slash at the end.
        sPath = Space$(MAX_PATH)
        idlstr = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        If idlstr Then
          GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1) & "\"
        End If
    End If

Exit_GetFolder:
    Exit Function
Err_GetFolder:
    MsgBox Err.Description, vbCritical Or vbOKOnly
    Resume Exit_GetFolder

End Function

''
' Ritorna l'estensione di un file
'
' @param Nome completo del file
Public Function GetFileExtension(ByVal FileName As String) As String
Dim pos As Integer

    pos = InStrRev(FileName, ".")
    If pos = 0 Then
        GetFileExtension = vbNullString
    Else
        GetFileExtension = Mid(FileName, pos)
    End If
End Function

''
' True se il file name è una directory
Public Function isDirectory(ByVal sFileName As String) As Boolean
Dim fd As WIN32_FIND_DATA
Dim h As Long

    h = FindFirstFile(sFileName, fd)
    isDirectory = (fd.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY)
    FindClose h
End Function

''
' Legge il valore di una chiave nel registry
'
' @param chiave radice
' @param nome della chiave sotto cui leggere
' @param nome del valore da leggere nella chiave
' @param valore letto
'
' @return True se tutto è andato bene
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef keyVal As String) As Boolean
On Error GoTo GetKeyError
Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim Hkey As Long                                        ' Handle To An Open Registry Key
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable

    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, 983103, Hkey) ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(Hkey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        keyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            keyVal = keyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        keyVal = Format$("&h" + keyVal)                     ' Convert Double Word To String
    End Select

    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(Hkey)                                  ' Close Registry Key
    Exit Function                                           ' Exit

GetKeyError:      ' Cleanup After An Error Has Occured...
    keyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(Hkey)                                  ' Close Registry Key
End Function

''
' WRITES A STRING VALUE TO REGISTRY:
'
' @param Top Level Key as defined by REG_TOPLEVEL_KEYS Enum (See Declarations)
' @param Full Path of Subkey if path does not exist it will be created
' @param ValueName
' @param Value Data
'
' @return: True if successful, false otherwise
Public Function WriteStringToRegistry(Hkey As Long, strPath As String, strValue As String, strdata As String) As Boolean
On Error GoTo ErrorHandler
Dim keyhand As Long
Dim r As Long

   'EXAMPLE: 'WriteStringToRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft", "CustomerName", "FreeVBCode.com")
   r = RegCreateKey(Hkey, strPath, keyhand)
   If r = 0 Then
        r = RegSetValueEx(keyhand, strValue, 0, _
           REG_SZ, ByVal strdata, Len(strdata))
        r = RegCloseKey(keyhand)
    End If
    
   WriteStringToRegistry = (r = 0)

Exit Function

ErrorHandler:
    WriteStringToRegistry = False
    Exit Function
End Function

'------------------------------------
'------------------------------------ PRIVATE
'------------------------------------

''
' Recupero del programma associato all'estensione
'
' @param nome dell'estensione
' @param percorso del programma associato
'
' @return true se tutto è ok
Private Function GetFileAssociation0(ByVal sExtName As String, ByRef sOutput As String) As Boolean
On Error GoTo err1
Dim Hkey      As Long
Dim strBuffer As String
Dim lResult   As Long
Dim lLen As Long
Dim strValueName As String
Dim lType As Long

    lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & sExtName & "\OpenWithList", 0, KEY_READ, Hkey)
    If lResult <> 0 Then
        sOutput = "Unregistered file extension " & sExtName
        GetFileAssociation0 = False
        Exit Function
    End If
    
    strValueName = "a"
    lType = REG_SZ
    strBuffer = Space$(128)
    lLen = Len(strBuffer)
    
    lResult = RegQueryValueEx(Hkey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
    
    If lResult <> 0 Then
        GetFileAssociation0 = False
        Exit Function
    End If
    
    sOutput = Mid$(strBuffer, 1, lLen - 1)
    GetFileAssociation0 = True
    Exit Function
    
err1:
    sOutput = Err.Description
    GetFileAssociation0 = False
End Function

''
' Recupero del programma associato all'estensione
'
' @param nome dell'estensione
' @param percorso del programma associato
'
' @return true se tutto è ok
Private Function GetFileAssociation1(ByVal sExtName As String, ByRef sOutput As String) As Boolean
On Error GoTo err1
Dim Hkey      As Long
Dim strBuffer As String
Dim lResult   As Long
Dim lLen As Long
Dim strValueName As String
Dim lType As Long

    lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & sExtName, 0, KEY_READ, Hkey)
    If lResult <> 0 Then
        sOutput = "Unregistered file extension " & sExtName
        GetFileAssociation1 = False
        Exit Function
    End If
    
    strValueName = "Application"
    lType = REG_SZ
    strBuffer = Space$(128)
    lLen = Len(strBuffer)
    
    lResult = RegQueryValueEx(Hkey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
    
    If lResult <> 0 Then
        GetFileAssociation1 = False
        Exit Function
    End If
    
    sOutput = Mid$(strBuffer, 1, lLen - 1)
    GetFileAssociation1 = True
    Exit Function
    
err1:
    sOutput = Err.Description
    GetFileAssociation1 = False
End Function

''
' Recupero del programma associato all'estensione
'
' @param nome dell'estensione
' @param percorso del programma associato
'
' @return true se tutto è ok
Private Function GetFileAssociation2(ByVal sExtName As String, ByRef sOutput As String) As Boolean
On Error GoTo err1
Dim Hkey      As Long
Dim strBuffer As String
Dim strTemp   As String
Dim lResult   As Long
Dim lLen As Long
Dim strValueName As String
Dim lType As Long

    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, sExtName, 0, KEY_READ, Hkey)
    If lResult <> 0 Then
        sOutput = "Unregistered file extension " & sExtName
        GetFileAssociation2 = False
        Exit Function
    End If
    
    strValueName = vbNullString
    lType = REG_SZ
    strBuffer = Space$(128)
    lLen = Len(strBuffer)
    
    lResult = RegQueryValueEx(Hkey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
    
    If lResult <> 0 Then
        GetFileAssociation2 = False
        Exit Function
    End If
    
    strTemp = Mid$(strBuffer, 1, lLen - 1)
    strTemp = strTemp & "\shell\open\command"
    
    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTemp, 0, KEY_READ, Hkey)
    If lResult <> 0 Then
        lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\" & strTemp, 0, KEY_READ, Hkey)
        If lResult <> 0 Then
            sOutput = "File type is not associated with a program."
            GetFileAssociation2 = False
            Exit Function
        End If
    End If
    
    lLen = Len(strBuffer)
    strValueName = vbNullString
    lResult = RegQueryValueEx(Hkey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
    sOutput = Mid$(strBuffer, 1, lLen - 1)
    
    GetFileAssociation2 = True
    Exit Function
    
err1:
    sOutput = Err.Description
    GetFileAssociation2 = False
End Function

''
' Get del path dal registry del programma
'
' @param nome del programma
'
' @return path completo del programma
Private Function GetProgramPath(ByVal sProgName As String) As String
On Error GoTo err1
Dim Hkey      As Long
Dim strBuffer As String
Dim strTemp   As String
Dim lResult   As Long
Dim lLen As Long
Dim strValueName As String
Dim lType As Long

    lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, "\Applications\" & sProgName & "\shell\open\command", 0, KEY_READ, Hkey)
    If lResult <> 0 Then
        GetProgramPath = sProgName
        Exit Function
    End If
    
    strValueName = vbNullString
    lType = REG_SZ
    strBuffer = Space$(128)
    lLen = Len(strBuffer)
    
    lResult = RegQueryValueEx(Hkey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
    
    If lResult <> 0 Then
        GetProgramPath = sProgName
        Exit Function
    End If
    
    strTemp = Mid$(strBuffer, 1, lLen - 1)
    strTemp = Replace(strTemp, """", "")
    GetProgramPath = Trim(Replace(strTemp, "%1", ""))
    
    If Dir(GetProgramPath) = vbNullString Then
        GetProgramPath = sProgName
    End If
    Exit Function
    
err1:
    GetProgramPath = sProgName
End Function

''
' Run program
Public Function RunProgram(ByVal sCommandLine As String, ByVal sCurrentDirectory As String, ByVal bWait As Boolean, _
                           ByVal sUserName As String, ByVal sPassword As String, ByVal sDomainName As String, _
                           ByVal bNETOnly As Boolean) As Boolean
Dim si As STARTUPINFO
Dim pi As PROCESS_INFORMATION
Dim wUser As String
Dim wDomain As String
Dim wPassword As String
Dim wCommandLine As String
Dim wCurrentDir As String
Dim lResult As Long
Dim pToken As Long
Dim lLogon As Long
    
    si.cb = Len(si)
        
    If sUserName <> "" Then
        wUser = StrConv(sUserName + Chr$(0), vbUnicode)
        wDomain = StrConv(sDomainName + Chr$(0), vbUnicode)
        wPassword = StrConv(sPassword + Chr$(0), vbUnicode)
        wCommandLine = StrConv(sCommandLine + Chr$(0), vbUnicode)
        wCurrentDir = StrConv(sCurrentDirectory + Chr$(0), vbUnicode)
    
        If bNETOnly Then
            lLogon = LOGON_NETCREDENTIALS_ONLY
        Else
            lLogon = LOGON_WITH_PROFILE
        End If
        
        lResult = CreateProcessWithLogonW(wUser, wDomain, wPassword, _
            lLogon, 0&, wCommandLine, _
            CREATE_DEFAULT_ERROR_MODE, 0&, wCurrentDir, si, pi)
                
        'alternativa da valutare eventualmente:
        'If LogonUser(sUserName, sDomainName, sPassword, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_WINNT50, pToken) <> 0 Then
        '   lResult = CreateProcessAsUser(pToken, 0&, sCommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS Or CREATE_SEPARATE_WOW_VDM, 0&, sCurrentDirectory, si, pi)
        'End If
    Else
        lResult = CreateProcess(0&, sCommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS Or CREATE_SEPARATE_WOW_VDM, 0&, sCurrentDirectory, si, pi)
    End If
    
    If lResult <> 0 Then
    
        If bWait Then
            Do
                lResult = WaitForSingleObject(pi.hProcess, 0)
                DoEvents
            Loop Until lResult <> 258
        End If
    
        CloseHandle pToken
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
        RunProgram = True
    Else
        RunProgram = False
        Debug.Assert False
        Debug.Print Err.LastDllError
    End If
End Function
