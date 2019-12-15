VERSION 5.00
Begin VB.UserControl ctlTable 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   2040
      Max             =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame frmBase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   2055
      Begin VB.TextBox txtCell 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "ctlTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event CellConfirm(ByVal iRowN As Integer, ByVal iColN As Integer, ByVal sText As String)

Private Const FIRST_CELL_TOP As Integer = 100
Private Const FIRST_CELL_LEFT As Integer = 10

Private iRow As Integer
Private iCol As Integer

Private Sub UserControl_Initialize()
    iRow = 1
    iCol = 1
    txtCell(0).Top = FIRST_CELL_TOP
    txtCell(0).Left = FIRST_CELL_LEFT
End Sub

Private Sub UserControl_Resize()
    frmBase.Width = UserControl.Width - VScroll1.Width
    frmBase.Height = UserControl.Height - frmBase.Top - HScroll1.Height
    
    VScroll1.Top = 0
    VScroll1.Left = frmBase.Width
    VScroll1.Height = frmBase.Height + frmBase.Top
    
    HScroll1.Top = frmBase.Top + frmBase.Height
    HScroll1.Left = frmBase.Left
    HScroll1.Width = frmBase.Width
End Sub

''
' Imposta il numero di celle
Public Sub setCellsNum(ByVal iNumRow As Integer, ByVal iNumCol As Integer)
Dim i As Integer
Dim iNumCell As Integer
Dim iTxtIndex As Integer

    iNumCell = iNumRow * iNumCol
    
    If txtCell.Count < iNumCell Then
        'espandere il numero di celle
        For i = txtCell.Count To iNumCell - 1
            Load txtCell(i)
        Next i
    ElseIf txtCell.Count > iNumCell Then
        'ridurre il numero di celle
        For i = txtCell.Count - 1 To iNumCell Step -1
            Unload txtCell(i)
        Next i
    End If
    
    iRow = iNumRow
    iCol = iNumCol
    
    VScroll1.Max = iRow
    HScroll1.Max = iCol
    
    ArrangeTxtCells
End Sub

''
' Imposta la dimensione delle celle
Public Sub setCellsDim(ByVal iWidth As Integer, ByVal iHeight As Integer)
Dim i As Integer

    For i = 0 To txtCell.Count - 1
        txtCell(i).Width = iWidth
        txtCell(i).Height = iHeight
    Next i
    ArrangeTxtCells
End Sub

''
' Imposta la dimensione delle celle di una colonna
Public Sub setColDim(ByVal iColN As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)
Dim i As Integer

    For i = 0 To txtCell.Count - 1
        If IsCellInCol(iColN, i) Then
            txtCell(i).Width = iWidth
            txtCell(i).Height = iHeight
        End If
    Next i
    ArrangeTxtCells
End Sub

''
' Imposta l'allinemaneto delle celle di una colonna
Public Sub setColAllignment(ByVal iColN As Integer, ByVal iAllign As Integer)
Dim i As Integer

    For i = 0 To txtCell.Count - 1
        If IsCellInCol(iColN, i) Then
            txtCell(i).Alignment = iAllign
        End If
    Next i
End Sub

''
' Imposta la dimensione delle celle di una riga
Public Sub setRowDim(ByVal iRowN As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer)
Dim i As Integer

    For i = iCol * (iRowN - 1) To (iCol * iRowN) - 1
        txtCell(i).Width = iWidth
        txtCell(i).Height = iHeight
    Next i
    ArrangeTxtCells
End Sub

''
' Imposta le scrollbar visibili
Public Sub setScrollbars(ByVal bVertical As Boolean, ByVal bHorizontal As Boolean)
Dim i As Integer

    HScroll1.Visible = bHorizontal
    HScroll1.value = HScroll1.Min
    VScroll1.Visible = bVertical
    VScroll1.value = VScroll1.Min
End Sub

''
' Imposta il testo in una cella
Public Sub setCellText(ByVal iRowN As Integer, ByVal iColN As Integer, ByVal sText As String)
Dim i As Integer

    i = convertCellNum(iRowN, iColN)
    If i >= 0 Then
        txtCell(i).Text = sText
    End If
End Sub

''
' Recupera il testo in una cella
Public Function getCellText(ByVal iRowN As Integer, ByVal iColN As Integer) As String
Dim i As Integer

    i = convertCellNum(iRowN, iColN)
    If i >= 0 Then
        getCellText = txtCell(i).Text
    End If
End Function

''
' Imposta il bocco su una cella
Public Sub setColLock(ByVal iColN As Integer, ByVal bLock As Boolean)
Dim i As Integer

    For i = 0 To txtCell.Count - 1
        If IsCellInCol(iColN, i) Then
            txtCell(i).Locked = bLock
            If bLock Then
                txtCell(i).BackColor = UserControl.BackColor
            Else
                txtCell(i).BackColor = vbWhite
            End If
        End If
    Next i
End Sub

'------------------------------------
'------------------------------------ PRIVATE
'------------------------------------

Private Sub ArrangeTxtCells()
Dim i As Integer
Dim j As Integer

    For i = 0 To txtCell.Count - 1
        If i = 0 Then
            txtCell(0).Top = -1 * txtCell(0).Height * VScroll1.value + FIRST_CELL_TOP
            txtCell(0).Left = -1 * txtCell(0).Width * HScroll1.value + FIRST_CELL_LEFT
        ElseIf (i Mod iCol) = 0 Then
            txtCell(i).Left = txtCell(0).Left
            txtCell(i).Top = txtCell(i - iCol).Top + txtCell(i - iCol).Height
        Else
            txtCell(i).Left = txtCell(i - 1).Left + txtCell(i - 1).Width
            txtCell(i).Top = txtCell(i - 1).Top
        End If
        
        txtCell(i).Visible = True
    Next i
End Sub

''
' Identifica l'indice di txtCell partendo dalla riga, colonna
Private Function convertCellNum(ByVal iRowN As Integer, ByVal iColN As Integer) As Integer
    convertCellNum = (iRowN - 1) * iCol + (iColN - 1)
End Function

''
' Identifica la riga, colonna partendo dall'indice di txtCell
Private Sub convertIndexNum(ByVal iIndex As Integer, ByRef iRowN As Integer, ByRef iColN As Integer)
    iRowN = Round(iIndex / iCol) + 1
    iColN = iIndex - ((iRowN * iCol) - iCol - 1)
End Sub

''
' Controlla se la cella appartiene alla colonna
Private Function IsCellInCol(ByVal iColN As Integer, ByVal iCell As Integer) As Boolean
    
    iCell = iCell + 1
    IsCellInCol = True
    
    If iCell = iColN Then Exit Function
    
    If iCol = iColN Then
        If (iCell Mod iCol) = 0 Then Exit Function
    Else
        If (iCell Mod iCol) = iColN Then Exit Function
    End If
    
    IsCellInCol = False
End Function

Private Sub VScroll1_Change()
    ArrangeTxtCells
End Sub

Private Sub HScroll1_Change()
    ArrangeTxtCells
End Sub

Private Sub txtCell_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo err1
Dim iNewIndex As Integer
    If KeyCode = vbKeyUp Then
        iNewIndex = Index - iCol
    ElseIf KeyCode = vbKeyDown Then
        iNewIndex = Index + iCol
    End If
    
    If iNewIndex > 0 And iNewIndex < txtCell.Count Then
        txtCell(iNewIndex).SetFocus
    End If
    Exit Sub
    
err1:
End Sub

Private Sub txtCell_KeyPress(Index As Integer, KeyAscii As Integer)
Dim iRowN As Integer
Dim iColN As Integer

    If KeyAscii = vbKeyReturn Then
        convertIndexNum Index, iRowN, iColN
        RaiseEvent CellConfirm(iRowN, iColN, txtCell(Index).Text)
    End If
End Sub

Private Sub txtCell_GotFocus(Index As Integer)
    If txtCell(Index).Locked Then
        If (Index + 1) + 1 <= txtCell.UBound Then
            If Not txtCell(Index + 1).Locked Then
                txtCell(Index + 1).SetFocus
            End If
        End If
    End If
End Sub
