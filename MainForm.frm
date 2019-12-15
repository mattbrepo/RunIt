VERSION 5.00
Begin VB.Form MainForm 
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4875
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdUsage 
      Caption         =   "Usage"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox chkWait 
      Caption         =   "Run and wait"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkRunIn 
      Caption         =   "Run in "
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame frameRunIn 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   4695
      Begin VB.TextBox txtDirectory 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDirectory 
         Caption         =   "Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CheckBox chkExit 
      Caption         =   "Run and exit"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkRunAs 
      Caption         =   "Run as"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame frameRunAs 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
      Begin VB.CheckBox chkNETOnly 
         Caption         =   "NET Only"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox txtGroup 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblGroup 
         Caption         =   "Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblUser 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdRun 
      Caption         =   "Run it"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblFileName 
      Caption         =   "Program file path (i.e.: c:\test.exe -v):"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mPopAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sBasePath As String

Private Sub cmdUsage_Click()
    '%cmdline%
    MsgBox "RunIt -f <filename> [-d <run dir>] [-u <username>] [-p <password>] [-g <group>] [-n <0;1>] [-w <0;1>] [-e <0;1>]" & vbNewLine & vbNewLine
End Sub

Private Sub Form_Load()
    Me.Caption = PROG_NAME & " v" & PROG_VER
    
    sBasePath = GetSpecialFolder(CSIDL_APPDATA) & FIRM_NAME & "\" & PROG_NAME & "\"
    If Dir(sBasePath, vbDirectory) = "" Then
        MkDir sBasePath
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub mPopAbout_Click()
    AboutForm.Show vbModal
End Sub

Private Sub chkRunAs_Click()
    frameRunAs.Enabled = (chkRunAs.value = vbChecked)
End Sub

Private Sub chkRunIn_Click()
    frameRunIn.Enabled = (chkRunIn.value = vbChecked)
End Sub

Public Sub CmdRun_Click()
Dim bWait As Boolean
Dim sProgramFile As String
Dim sDirectory As String
Dim sUser As String
Dim sPassword As String
Dim sGroup As String
Dim bNETOnly As Boolean
Dim bResult As Boolean

    Me.Visible = (chkExit.value = vbUnchecked)
    
    bWait = (chkWait.value = vbChecked)
    sProgramFile = txtFileName.Text
    sDirectory = txtDirectory.Text
    
    If chkRunAs.value = vbChecked Then
        sUser = txtUser.Text
        sPassword = txtPassword.Text
        sGroup = txtGroup.Text
        bNETOnly = (chkNETOnly.value = vbChecked)
    Else
        sUser = ""
        sPassword = ""
        sGroup = ""
        bNETOnly = False
    End If
    
    bResult = RunProgram(sProgramFile, sDirectory, bWait, sUser, sPassword, sGroup, bNETOnly)
    
    If (chkExit.value = vbChecked) Then
        Unload Me
    ElseIf Not bResult Then
        MsgBox "Run program error: " & sProgramFile
    End If
End Sub

Private Sub txtFileName_Change()
On Error GoTo err1
Dim iPos1 As Integer
Dim iPos2 As Integer
    
    If chkRunIn.value = vbChecked Then Exit Sub

    iPos1 = InStr(txtFileName.Text, ".exe")
    If iPos1 > 0 Then
        iPos2 = InStrRev(Mid(txtFileName.Text, 1, iPos1), "\")
        txtDirectory.Text = Mid(txtFileName.Text, 1, iPos2)
    Else
        txtDirectory.Text = ""
    End If
    Exit Sub
err1:
    txtDirectory.Text = ""
End Sub
