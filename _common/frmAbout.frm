VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLink 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "http://www.test.com"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txtEmail 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "test@test.com"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtAbout 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   154
      TabIndex        =   3
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label lblLinkAct 
      Alignment       =   2  'Center
      Caption         =   "http://www.test.com/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblEmailAct 
      Alignment       =   2  'Center
      Caption         =   "test@test.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblLink 
      Caption         =   "Link:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set Me.Icon = MainForm.Icon
    txtAbout.Text = PROG_NAME & " v" & PROG_VER
    
    txtEmail.BackColor = Me.BackColor
    txtLink.BackColor = Me.BackColor
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then cmdOk_Click
End Sub
