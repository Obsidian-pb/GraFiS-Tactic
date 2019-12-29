VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_PressureChange 
   Caption         =   "Напоры на входе"
   ClientHeight    =   4950
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   3864
   OleObjectBlob   =   "f_PressureChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_PressureChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public currentShp As Visio.Shape


Private Sub CB_Cancel_Click()
    Me.Hide
End Sub

Private Sub CB_OK_Click()
    
'---Обновляем данные по напорам
    RefreshTBData Me.TB_InPatr1, GetCellName("GFS_InPatrIn1")
    RefreshTBData Me.TB_InPatr2, GetCellName("GFS_InPatrIn2")
    RefreshTBData Me.TB_InWCLeft1, GetCellName("GFS_InWCLeft1")
    RefreshTBData Me.TB_InWCLeft2, GetCellName("GFS_InWCLeft2")
    RefreshTBData Me.TB_InWCRight1, GetCellName("GFS_InWCRight1")
    RefreshTBData Me.TB_InWCRight2, GetCellName("GFS_InWCRight2")
    RefreshTBData Me.TB_Tank1, GetCellName("GFS_InTank1")
    RefreshTBData Me.TB_Tank2, GetCellName("GFS_InTank2")

'---Закрываем форму
    Me.Hide
End Sub

Private Sub UserForm_Activate()

    CB_OK.SetFocus

    SetTBData Me.TB_InPatr1, GetCellName("GFS_InPatrIn1")
    SetTBData Me.TB_InPatr2, GetCellName("GFS_InPatrIn2")
    SetTBData Me.TB_InWCLeft1, GetCellName("GFS_InWCLeft1")
    SetTBData Me.TB_InWCLeft2, GetCellName("GFS_InWCLeft2")
    SetTBData Me.TB_InWCRight1, GetCellName("GFS_InWCRight1")
    SetTBData Me.TB_InWCRight2, GetCellName("GFS_InWCRight2")
    SetTBData Me.TB_Tank1, GetCellName("GFS_InTank1")
    SetTBData Me.TB_Tank2, GetCellName("GFS_InTank2")
    
End Sub



Private Sub SetTBData(ByRef tb As textBox, ByVal cellName As String)
'Прока ставит для текстового поля значение содержащееся в указанной ячейке текущией фигуры
'Если такой ячейки нет - поле блокируется
    On Error GoTo EX
    
    tb.Value = currentShp.Cells(cellName).Result(visNumber)
    tb.Visible = True
    
Exit Sub
EX:
    tb.Value = ""
    tb.Visible = False
End Sub

Private Sub RefreshTBData(ByRef tb As textBox, ByVal cellName As String)
'Прока обновляет данные в секции Scratch согласно указанным названиям ячеек
    On Error GoTo EX
    
    currentShp.Cells(cellName).FormulaU = Chr(34) & tb.Value & Chr(34)

EX:
End Sub

Private Function GetCellName(ByVal connecionCellName) As String
'Функция возвращает имя строки по указанному имени строки в секции Connections
Dim i As Integer
Dim cll As Visio.cell

    For i = 0 To currentShp.RowCount(visSectionConnectionPts) - 1
        If connecionCellName = currentShp.CellsSRC(visSectionConnectionPts, i, 0).RowNameU Then
            GetCellName = "Scratch.C" & i + 1
            Exit Function
        End If
    Next i
    
EX:
End Function








