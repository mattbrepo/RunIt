Attribute VB_Name = "mdlMain"
Option Explicit

Public Const FIRM_NAME As String = "SofaUtil"
Public sConfigFile As String

Public Sub Main()
Dim sConfigDir As String

    sConfigDir = GetSpecialFolder(CSIDL_APPDATA) & FIRM_NAME & "\"
    
    If Dir(sConfigDir, vbDirectory) = "" Then
        MkDir sConfigDir
    End If
    
    sConfigFile = sConfigDir & mdlMain0.PROG_NAME & "_config.ini"
    
    mdlConfig.LoadConfig sConfigFile
    
    mdlMain0.Run Command()
End Sub
