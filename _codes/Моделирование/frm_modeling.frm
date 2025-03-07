VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_modeling 
   Caption         =   "Моделирование площади"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   OleObjectBlob   =   "frm_modeling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_modeling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fireShape As Visio.Shape            'Фигура площади пожара

Dim matrixSize As Long              'Количество клеток в матрице
Dim matrixChecked As Long           'Количество проверенных клеток





Private Sub B_Cancel2_Click()
    Me.Hide
End Sub

Private Sub cb_run_Click()
    GRAIN = Int(Me.tb_grain)
    MakeMatrixM Me, Me.tb_rounds, Me.tb_calcTime, Me.tb_checkTime
End Sub


'--------------------------Внутрение процедуры МОДЕЛИРОВАНИЕ----------------------------------
Private Function GetMatrixCheckedStatus(Optional kind As Byte = 0) As String
'Возвращает подпись для статуса запекания матрицы
'Dim procent As Single
'    procent = Round(matrixChecked / matrixSize, 4) * 100
'
'    GetMatrixCheckedStatus = "Запечено " & procent & "%"
    
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    Select Case kind
        Case Is = 0
            GetMatrixCheckedStatus = "Запечено " & procent & "%"
        Case Is = 1
            GetMatrixCheckedStatus = "Обработка расчетной зоны " & procent & "%"
        Case Is = 2
            GetMatrixCheckedStatus = "Обработано " & procent & "% стен"
        Case Is = 3
            GetMatrixCheckedStatus = "Обработано " & procent & "% дверей"
    End Select
End Function


'--------------------------Внешние процедуры и функции МОДЕЛИРОВАНИЕ--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'Указываем для формы общее кол-во клеток в матрице
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long, Optional kind As Byte = 0)
'Добавляем кол-во проверенных клеток
    matrixChecked = matrixChecked + size
    
    'Обновляем статусную строку с количеством проверенных клеток
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus(kind)
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub
