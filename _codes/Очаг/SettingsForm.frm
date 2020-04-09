VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Параметры построения"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   OleObjectBlob   =   "SettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim matrixSize As Long          'Количество клеток в матрице
Dim matrixChecked As Long       'Количество проверенных клеток



Private Sub btnStopModelling_Click()
    stopModellingFlag = True
End Sub

Private Sub optTTX_Change()
'    txtNozzleRangeValue.Enabled = False
End Sub



Private Sub optValue_Change()
    txtNozzleRangeValue.Enabled = optValue.value
    If txtNozzleRangeValue.Enabled Then
        txtNozzleRangeValue.BackColor = vbWhite
    Else
        txtNozzleRangeValue.BackColor = &H8000000F
    End If
    
End Sub



'------------------------Процедуры, собственно формы--------------------------
Private Sub UserForm_Activate()
    Me.txtGrainSize = grain
    matrixSize = 0
    matrixChecked = 0
    
    'Проверяем, запечена ли матрица
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "Матрица запечена. Размер зерна " & grain & "мм."
        lblMatrixIsBaked.ForeColor = vbGreen
    Else
        lblMatrixIsBaked.Caption = "Матрица не запечена."
        lblMatrixIsBaked.ForeColor = vbRed
    End If
    
    txtGrainSize.value = 200
End Sub



Private Sub btnBakeMatrix_Click()
    'Запоминаем значение зерна матрицы
    grain = Me.txtGrainSize
    
    'Запекаем матрицу
    MakeMatrix
End Sub

Private Sub btnDeleteMatrix_Click()
    'Удаляем матрицу
    DestroyMatrix
    
    'Указываем, что матрица не запечена
    lblMatrixIsBaked.Caption = "Матрица не запечена."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRunFireModelling_Click()
'При нажатии на кнопку запускаем моделирование
    stopModellingFlag = False
    
    On Error GoTo EX
    'Определяем требуемое количество шагов
    Dim spd As Single
    Dim timeElapsed As Single
    Dim intenseNeed As Single
    spd = Me.txtSpeed
    timeElapsed = Me.txtTime
    intenseNeed = CSng(Replace(Me.txtIntense, ".", ","))
    
    'проверяем, все ли данные указаны верно
    If timeElapsed > 0 And spd > 0 Then
        'Строим площадь
        RunFire timeElapsed, spd, intenseNeed
    Else
        MsgBox "Не все данные корректно указаны!", vbCritical
    End If
Exit Sub
EX:
    MsgBox "Не все данные корректно указаны!", vbCritical
End Sub




'--------------------------внутрение процедуры----------------------------------
Private Function GetMatrixCheckedStatus() As String
'Возвращает подпись для статуса запекания матрицы
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    GetMatrixCheckedStatus = "Запечено " & procent & "%"
End Function





'--------------------------Внешние процедуры и функции--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'Указываем для формы общее кол-во клеток в матрице
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long)
'Добавляем кол-во проверенных клеток
    matrixChecked = matrixChecked + size
    
    'Обновляем статусную строку с количество проверенных клеток
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub






























