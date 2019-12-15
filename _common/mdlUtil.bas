Attribute VB_Name = "mdlUtil"
Option Explicit

''
' Ritorna se l'array è vuoto
'
' @param array
Public Function IsEmptyArray(ByRef parArr) As Boolean
On Error GoTo err1
    
    IsEmptyArray = True
    
    If UBound(parArr) > 0 Then IsEmptyArray = False
err1:
End Function

''
' Tokenizza una stringa
Public Function GetToken(ByRef sStr As String, ByVal sDivisore As String) As Variant
    If (InStr(1, sStr, sDivisore) - 1 >= 0) Then
        GetToken = Left$(sStr, InStr(1, sStr, sDivisore) - 1)
        sStr = Mid$(sStr, InStr(1, sStr, sDivisore) + 1)
    Else
        GetToken = sStr
        sStr = ""
        Exit Function
    End If
End Function

