VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InsertExtSquare 
   Caption         =   "Рассчитать площадь тушения"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   OleObjectBlob   =   "F_InsertExtSquare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InsertExtSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fireShape As Visio.Shape            'Фигура площади пожара



'Public Vfl_TargetShapeID As Long
'Public VmD_TimeStart As Date
'Private vfStr_ObjList() As String


Dim matrixSize As Long              'Количество клеток в матрице
Dim matrixChecked As Long           'Количество проверенных клеток
'Public timeElapsedMain As Single    'Время прошедшее с начала моделирования
'Public pathMain As Single           'Пройденный путь с начала моделирования



'--------------------------------Блок основных процедур----------------------------------------------------------------

Private Sub UserForm_Initialize()
'Процедура загрузки формы

    '---Значения выпадающих списков
    FillCBCalculateType

End Sub



Public Function SetFireShape(ByRef shp As Visio.Shape) As F_InsertExtSquare
    Set fireShape = shp
    Set SetFireShape = Me
    
    Me.Caption = "Расчитать площадь тушения (" & shp.Name & ")"
End Function

Private Sub UserForm_Activate()
'Процедура активации формы - при показе
    
    'Проверяем, запечена ли матрица
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "Матрица запечена. Размер зерна " & grain & "мм."
        lblMatrixIsBaked.ForeColor = vbGreen
        Me.txtGrainSize = grain
    Else
        lblMatrixIsBaked.Caption = "Матрица не запечена."
        lblMatrixIsBaked.ForeColor = vbRed
        Me.txtGrainSize.value = 200
    End If
    
End Sub

Private Sub btnRunExtSquareCalc_Click()
'При нажатии на кнопку запускаем расчет и построение площади
Dim extSquareCalculator As c_ExtSquareCalculator

    'Проверяем, запечена ли матрица
    If Not IsMatrixBacked Then
        MsgBox "Матрица не запечена!!!"
        Exit Sub
    End If
    
    'Собственно запускаем расет и построение
    Set extSquareCalculator = New c_ExtSquareCalculator
    extSquareCalculator.SetOpenSpaceLayer fireModeller
    extSquareCalculator.RunDemon fireShape
    
    Me.Hide
End Sub

Private Sub B_Cancel2_Click()
    Me.Hide
End Sub

Private Sub btnBakeMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "Слишком большой размер результирующей матрицы! Уменьшите размер рабочего листа или зерна матрицы.", vbInformation, "ГраФиС-Тактик"
        Exit Sub
    End If

    'Запоминаем значение зерна матрицы
    grain = Me.txtGrainSize
    
    'Запекаем матрицу
    MakeMatrix Me
End Sub

Private Sub btnDeleteMatrix_Click()
    'Удаляем матрицу
    DestroyMatrix
        
    'Указываем, что матрица не запечена
    lblMatrixIsBaked.Caption = "Матрица не запечена."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRefreshMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "Слишком большой размер результирующей матрицы! Уменьшите размер рабочего листа или зерна матрицы.", vbInformation, "ГраФиС-Тактик"
        Exit Sub
    End If
    
    'Обновляем матрицу открытых пространств
    RefreshOpenSpacesMatrix Me
End Sub

'--------------------------Внутрение процедуры МОДЕЛИРОВАНИЕ----------------------------------
Private Function GetMatrixCheckedStatus() As String
'Возвращает подпись для статуса запекания матрицы
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    GetMatrixCheckedStatus = "Запечено " & procent & "%"
End Function


'--------------------------Внешние процедуры и функции МОДЕЛИРОВАНИЕ--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'Указываем для формы общее кол-во клеток в матрице
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long)
'Добавляем кол-во проверенных клеток
    matrixChecked = matrixChecked + size
    
    'Обновляем статусную строку с количеством проверенных клеток
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub


'------------------------Наполнение списков
Private Sub FillCBCalculateType()
'Список возможных вариантов расчета
    CB_CalculateType.AddItem "Требуемая площадь тушения"
    CB_CalculateType.AddItem "Фактическая площадь тушения"
    CB_CalculateType.AddItem "Весь периметр"
    CB_CalculateType.ListIndex = 0
    CB_CalculateType.ControlTipText = "Весь периметр - весь периметр фигуры рассматривается как фронт пожара, не зависимо от конфигурации ограждающих конструкций;" & Chr(13) & Chr(10) & _
    "Требуемая площадь тушения - расчет проводится только для участков фронта пожара, участки, граничащие с ограждающими конструкциями, не рассматриваются;" & Chr(13) & Chr(10) & _
    "Фактическая площадь тушения - рассчитываются только те участки фронта пожара, на которых поданы стволы"
End Sub

Public Property Get AttackDeep() As Byte
    AttackDeep = 5
End Property


