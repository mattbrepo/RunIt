Attribute VB_Name = "mdlMain0"
Option Explicit

Public Const PROG_NAME As String = "RunIt"
Public Const PROG_VER As String = "0.3"

Public Sub Run(ByVal sCommand As String)
Dim sVal As String
Dim sPar As String

    If sCommand = "" Then
        MainForm.Show
    Else
        Load MainForm
        
        'gestione della command line (vedi usage: %cmdline%)
        sPar = ""
        Do
            sVal = GetToken(sCommand, " ")
            If Left(sVal, 1) = "-" Then
                sPar = sVal
            
                If Left(sCommand, 1) = """" Then
                    sCommand = Mid(sCommand, 2)
                    sVal = GetToken(sCommand, """")
                Else
                    sVal = GetToken(sCommand, " ")
                End If
                
                Select Case LCase(sPar)
                Case "-f"
                    MainForm.txtFileName.Text = sVal
                Case "-d"
                    MainForm.chkRunIn.value = vbChecked
                    MainForm.txtDirectory.Text = sVal
                Case "-u"
                    MainForm.chkRunAs.value = vbChecked
                    MainForm.txtUser.Text = sVal
                Case "-p"
                    MainForm.txtPassword.Text = sVal
                Case "-g"
                    MainForm.txtGroup.Text = sVal
                Case "-w"
                    MainForm.chkWait.value = IIf((sVal = "1"), 1, 0)
                Case "-e"
                    MainForm.chkExit.value = IIf((sVal = "1"), 1, 0)
                Case "-n"
                    MainForm.chkNETOnly.value = IIf((sVal = "1"), 1, 0)
                End Select
            End If
        Loop While sCommand <> ""
        
        MainForm.CmdRun_Click
    End If
End Sub
