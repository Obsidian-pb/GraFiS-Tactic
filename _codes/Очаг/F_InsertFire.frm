VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InsertFire 
   Caption         =   "Укажите исходные данные"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   OleObjectBlob   =   "F_InsertFire.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InsertFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Vfl_TargetShapeID As Long
Public VmD_TimeStart As Date
Private VfB_NotShowPropertiesWindow As Boolean
Private vfStr_ObjList() As String


Dim matrixSize As Long              'Количество клеток в матрице
Dim matrixChecked As Long           'Количество проверенных клеток
Public timeElapsedMain As Single    'Время прошедшее с начала моделирования
Public pathMain As Single           'Пройденный путь с начала моделирования


'--------------------------------Блок ПОСТРОЕНИЕ---------------------------------------------------
Private Sub B_Cancel_Click()
'---Включаем показ окон
    VfB_NotShowPropertiesWindow = False
    
    Me.Hide
End Sub




Private Sub B_OK_Click()

'---Проверяем корректность указанных пользователем данных
    If fC_DataCheck = False Then Exit Sub

    Me.Hide
    
'---Проверяем какая опция построения площади выбрана и в зависимости от этого
    'вбрасываем готовую форму площади или прогнозируем
    If Me.OB_SolidShape.value = True Then
        s_FireShapeDrop
    ElseIf Me.OB_LiquidShape.value = True = True Then
        s_PrognoseFire
    End If
    
'---Включаем показ окон
    VfB_NotShowPropertiesWindow = False
End Sub

Private Sub B_Test_Click()
    s_PrognoseFire
End Sub






'--------------------------------Блок МОДЕЛИРОВАНИЕ---------------------------------------------------
Private Sub B_Cancel2_Click()
'---Включаем показ окон
    VfB_NotShowPropertiesWindow = False
    
    Me.Hide
End Sub

Private Sub btnBakeMatrix_Click()
    If IsAcceptableMatrixSize(1200000) = False Then
        MsgBox "Слишком большой размер результирующей матрицы! Уменьшите размер рабочего листа или зерна матрицы."
        Exit Sub
    End If

    'Запоминаем значение зерна матрицы
    grain = Me.txtGrainSize
    
    'Запекаем матрицу
    MakeMatrix
End Sub

Private Sub btnDeleteMatrix_Click()
    'Удаляем матрицу
    DestroyMatrix
    
    'Удаляем фигуры слоя Fire
    ClearLayer "Fire"
    
    'Указываем, что матрица не запечена
    lblMatrixIsBaked.Caption = "Матрица не запечена."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRunFireModelling_Click()
'При нажатии на кнопку запускаем моделирование
    'Проверяем, запечена ли матрица
    If Not IsMatrixBacked Then
        MsgBox "Матрица не запечена!!!"
        Exit Sub
    End If
    
    stopModellingFlag = False
    
    On Error GoTo EX
    'Определяем требуемое количество шагов
    Dim spd As Single
    Dim timeElapsed As Single
    Dim intenseNeed As Single
    
    'Определяем линейную скорость
    spd = GetSpeed
    'Определяем время моделирвоания
    '---Определяем время моделирования в соответствии с указанными пользователем значениями
    Dim vsStr_FirePath As String
    Dim vsD_TimeCur As Date
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        timeElapsed = DateDiff("s", VmD_TimeStart, vsD_TimeCur) / 60 ' В МИНУТАХ!!!!!!
    End If
    If Me.OB_ByDuration = True Then
        timeElapsed = Me.TB_Duration
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = str(ffSng_PointChange(Me.TB_Radius) * 2) & "m"
        If timeElapsedMain > 10 Then
            timeElapsed = CSng(ffSng_PointChange(Me.TB_Radius) / spd)      ' В МИНУТАХ!!!!!!
        Else
            timeElapsed = CSng(ffSng_PointChange(Me.TB_Radius) / (spd / 2))    ' В МИНУТАХ!!!!!!
        End If
    End If
    
    'Определяем требуемую интенсивность
    intenseNeed = GetIntense
    
    'проверяем, все ли данные указаны верно
    If timeElapsed > 0 And spd > 0 Then
        'Строим площадь
        If Me.OB_ByRadius = True Then       'Если моделирвоание осущесмтвляется по радиусу...
            RunFire timeElapsed, spd, intenseNeed, CSng(ffSng_PointChange(Me.TB_Radius))
        Else
            RunFire timeElapsed, spd, intenseNeed
        End If
        
    Else
        MsgBox "Не все данные корректно указаны!", vbCritical
        Exit Sub
    End If
    
    '---Передаем вброшенной фигуре значение объекта пожара из формы добавления площади
    Dim vsO_FireShape As Visio.Shape
    Set vsO_FireShape = Application.ActiveWindow.Selection(1)
    If Me.OB_SpeedByObject.value = True Then
        vsO_FireShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_FireShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_FireShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
    End If
    '---Добавляем вычисленную дату/время для полученной фигуры площади горения
    Dim actTime As Date
    actTime = DateAdd("n", timeElapsedMain, VmD_TimeStart)
    vsO_FireShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(actTime)) & ")"
    
Exit Sub
EX:
    MsgBox "Не все данные корректно указаны!", vbCritical
End Sub

Private Sub btnStopModelling_Click()
'Делаем паузу в моделировании
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
    
    'Обновляем статусную строку с количество проверенных клеток
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub







'--------------------------------Блок выбора типа рассчета пути---------------------------------------------------
Private Sub OB_ByDuration_AfterUpdate()
    Me.TB_Duration.Enabled = True
    Me.TB_Radius.Enabled = False
    Me.TB_Time.Enabled = False
    Me.TB_Duration.BackStyle = fmBackStyleOpaque
    Me.TB_Radius.BackStyle = fmBackStyleTransparent
    Me.TB_Time.BackStyle = fmBackStyleTransparent
    
    
End Sub


Private Sub OB_ByRadius_AfterUpdate()
    Me.TB_Radius.Enabled = True
    Me.TB_Duration.Enabled = False
    Me.TB_Time.Enabled = False
    Me.TB_Radius.BackStyle = fmBackStyleOpaque
    Me.TB_Duration.BackStyle = fmBackStyleTransparent
    Me.TB_Time.BackStyle = fmBackStyleTransparent
    
    
End Sub


Private Sub OB_ByTime_AfterUpdate()
    Me.TB_Time.Enabled = True
    Me.TB_Duration.Enabled = False
    Me.TB_Radius.Enabled = False
    Me.TB_Time.BackStyle = fmBackStyleOpaque
    Me.TB_Duration.BackStyle = fmBackStyleTransparent
    Me.TB_Radius.BackStyle = fmBackStyleTransparent
    
    
End Sub


'--------------------------------Блок выбора скорости----------------------------------------------------------------
Private Sub OB_SpeedByObject_AfterUpdate()
    Me.CB_Object.Enabled = True
    Me.CB_ObjectType.Enabled = True
    Me.TB_Speed2.Enabled = False
    Me.TB_Intense2.Enabled = False
    Me.CB_Object.BackStyle = fmBackStyleOpaque
    Me.CB_ObjectType.BackStyle = fmBackStyleOpaque
    Me.TB_Speed2.BackStyle = fmBackStyleTransparent
    Me.TB_Speed1.BackStyle = fmBackStyleOpaque
    Me.TB_Intense2.BackStyle = fmBackStyleTransparent
    Me.TB_Intense1.BackStyle = fmBackStyleOpaque

End Sub


Private Sub OB_SpeedByDirect_AfterUpdate()
    Me.CB_Object.Enabled = False
    Me.CB_ObjectType.Enabled = False
    Me.TB_Speed2.Enabled = True
    Me.TB_Intense2.Enabled = True
    Me.CB_Object.BackStyle = fmBackStyleTransparent
    Me.CB_ObjectType.BackStyle = fmBackStyleTransparent
    Me.TB_Speed2.BackStyle = fmBackStyleOpaque
    Me.TB_Speed1.BackStyle = fmBackStyleTransparent
    Me.TB_Intense2.BackStyle = fmBackStyleOpaque
    Me.TB_Intense1.BackStyle = fmBackStyleTransparent

End Sub


'--------------------------------Блок выбора типа построения-----------------------------------------------------------
Private Sub OB_SolidShape_AfterUpdate()
    Me.TB_REI.Enabled = False
    Me.TB_REI.BackStyle = fmBackStyleTransparent
    Me.CB_CheckOpens.Enabled = False
    Me.CB_CheckOpens.BackStyle = fmBackStyleTransparent
    
    Me.CB_Shape.Enabled = True
    Me.CB_Shape.BackStyle = fmBackStyleOpaque
End Sub

Private Sub OB_LiquidShape_AfterUpdate()
    Me.TB_REI.Enabled = True
    Me.TB_REI.BackStyle = fmBackStyleOpaque
    Me.CB_CheckOpens.Enabled = True
    Me.CB_CheckOpens.BackStyle = fmBackStyleOpaque
    
    Me.CB_Shape.Enabled = False
    Me.CB_Shape.BackStyle = fmBackStyleTransparent
End Sub



'--------------------------------Блок основных процедур----------------------------------------------------------------

Private Sub UserForm_Initialize()
'Процедура загрузки формы

'---Обновляем списки:
    sf_ObjectsListCreation 'Главный список объектов и скоростей
    
    '---Значения выпадающих списков
    sf_ObjectTypesListRefresh
    sf_ObjectsListRefresh
    sf_FireFormLoad
    
End Sub


Private Sub UserForm_Activate()
'Процедура активации формы - при показе
'---Отключаем показ окон
    VfB_NotShowPropertiesWindow = True

'---Переводим значение радиуса в значение с запятой
'    Me.TB_Radius.Value = ffSng_PointChange(Me.TB_Radius.Value)  ОБЯЗАТЕЛЬНО РАЗОБРАТЬСЯ!!!!

'---Активируем настройки МОДЕЛИРОВАНИЯ
    matrixSize = 0
    matrixChecked = 0
    
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

Private Sub sf_ObjectTypesListRefresh()
'Процедура обновления списка категорий объектов
Dim vsO_DBS As DAO.Database, vsO_RST As DAO.Recordset
Dim vsStr_SQL As String
Dim vsStr_Pth As String

    On Error GoTo EX
'---Очищаем имеющиеся списки
'    If CB_ObjectType.ListCount > 0 Then Exit Sub 'В случае, если список уже заполнен - не обновляем его
    Me.CB_ObjectType.Clear
'    Me.CB_Object.Clear

'---Определяем запрос SQL для отбора записей категорий из базы данных
    vsStr_SQL = "SELECT КатегорииОбъектов.Категория FROM КатегорииОбъектов;"
    
'---Создаем набор записей для получения списка категорий
    vsStr_Pth = ThisDocument.path & "Signs.fdb"
    Set vsO_DBS = GetDBEngine.OpenDatabase(vsStr_Pth)
    Set vsO_RST = vsO_DBS.CreateQueryDef("", vsStr_SQL).OpenRecordset(dbOpenDynaset)  'Создание набора записей


'---Ищем необходимую запись в наборе данных и добавляем её в список категорий объектов
    With vsO_RST
        .MoveFirst
        Do Until .EOF
            Me.CB_ObjectType.AddItem !Категория
            .MoveNext
        Loop
    End With

'---Активируем первую позицию списка
    Me.CB_ObjectType.ListIndex = 0

'---Очищаем объекты
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
Exit Sub
EX:
'---Очищаем объекты
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
    SaveLog Err, "sf_ObjectTypesListRefresh"
End Sub


Private Sub sf_ObjectsListCreation()
'Процедура формирует список объектов пожара
Dim vsO_DBS As DAO.Database, vsO_RST As DAO.Recordset
Dim vsStr_SQL As String
Dim vsStr_Pth As String
Dim i As Integer

    On Error GoTo EX
'---Определяем запрос SQL для отбора записей категорий из базы данных
    vsStr_SQL = "SELECT Категория, Описание, СкоростьРасч, ИнтенсивностьПоВодеРасч FROM З_Интенсивности;"
    
'---Создаем набор записей для получения списка категорий
    vsStr_Pth = ThisDocument.path & "Signs.fdb"
    Set vsO_DBS = GetDBEngine.OpenDatabase(vsStr_Pth)
    Set vsO_RST = vsO_DBS.CreateQueryDef("", vsStr_SQL).OpenRecordset(dbOpenDynaset)  'Создание набора записей

'---Ищем необходимую запись в наборе данных и добавляем её в список категорий объектов
    With vsO_RST
    i = 0
    '---Переобъявляем двумерный массив
        vsO_RST.MoveLast
        ReDim vfStr_ObjList(vsO_RST.RecordCount, 4) As String
    '---Обновляем список обюъектов пожара и их скоростей
        .MoveFirst
        Do Until .EOF
'            If !Категория = Me.CB_ObjectType.Value Then
            vfStr_ObjList(i, 0) = !Категория
            vfStr_ObjList(i, 1) = !Описание
            If !СкоростьРасч >= 0 Then vfStr_ObjList(i, 2) = !СкоростьРасч Else vfStr_ObjList(i, 2) = 0                             'Скорость распространения, линейная
            If !ИнтенсивностьПоВодеРасч >= 0 Then vfStr_ObjList(i, 3) = !ИнтенсивностьПоВодеРасч Else vfStr_ObjList(i, 3) = "0.1"   'Интенсивность по воде
'            Debug.Print !Описание & " скор:" & vfStr_ObjList(i, 2) & " инт:" & vfStr_ObjList(i, 3)
            i = i + 1
'            End If
            .MoveNext
        Loop
    End With

'---Очищаем объекты
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
Exit Sub
EX:
    SaveLog Err, "sf_ObjectsListCreation"
End Sub

Private Sub sf_ObjectsListRefresh()
'Процедура обновления списка объектов пожара
Dim i As Integer

Me.CB_Object.Clear

For i = 0 To UBound(vfStr_ObjList()) - 1
    If vfStr_ObjList(i, 0) = Me.CB_ObjectType.value Then
        Me.CB_Object.AddItem vfStr_ObjList(i, 1)
    End If
Next i

Me.CB_Object.ListIndex = 0

End Sub




Private Sub CB_ObjectType_Change()
'Обновляем список объектов в соответствии с новым значением категории
    sf_ObjectsListRefresh
End Sub

Private Sub CB_Object_Change()
'Получаем значение скорости в зависимости от указанного объекта пожара
Dim i As Integer

    For i = 0 To UBound(vfStr_ObjList()) - 1
        If vfStr_ObjList(i, 1) = Me.CB_Object.value Then
            Me.TB_Speed1.value = vfStr_ObjList(i, 2)
            Me.TB_Intense1.value = vfStr_ObjList(i, 3)
        End If
    Next i

End Sub

Private Sub sf_FireFormLoad()
'Определяем список форм пожара

    With Me.CB_Shape
        .AddItem "Площадь прямоугольная"
        .AddItem "Площадь круглая"
        .AddItem "Сектор 90"
        .AddItem "Сектор 180"
        .AddItem "Сектор 270"
        .ListIndex = 0
    End With

End Sub


'-------------------------------------------Блок вброса фигуры площади горения---------------------------------------

Private Sub s_FireShapeDrop()
'Процедура вброса фигуры площади и определения её стартовых параметров
Dim vsO_DropMaster As Visio.Master
Dim vsO_DropShape As Visio.Shape
Dim vsO_DropTargetShape As Visio.Shape
Dim vss_Speed As Single
Dim vsStr_FirePath As String
Dim vsD_TimeCur As Date
Dim vsL_Duration As Long

'---Активируем обработку ошибок
    On Error GoTo EX
'---Определяем рабочие объекты
    Set vsO_DropTargetShape = Application.ActivePage.Shapes.ItemFromID(Vfl_TargetShapeID)
    Set vsO_DropMaster = ThisDocument.Masters(Me.CB_Shape.value)

'---Вбрасываем фигуру площади горения и определяем её месторасположения
    Set vsO_DropShape = Application.ActivePage.Drop(vsO_DropMaster, 0, 0)
    vsO_DropShape.Cells("PinX").FormulaU = vsO_DropTargetShape.Cells("PinX").FormulaU
    vsO_DropShape.Cells("PinY").FormulaU = vsO_DropTargetShape.Cells("PinY").FormulaU
    vsO_DropShape.SendToBack

'---Определяем рассчетную скорость распространения огня
    vss_Speed = GetSpeed
'    If Me.OB_SpeedByDirect = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed2))
'    ElseIf Me.OB_SpeedByObject = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed1))
'    End If
    
'---Определяем размеры фигуры в соответствии с путем пройденным огнем в соответствии _
    с указанными пользователем значениями
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        vsL_Duration = DateDiff("s", VmD_TimeStart, vsD_TimeCur) ' В СЕКУНДАХ!!!!!!
        If vsL_Duration / 60 > 10 Then
            vsStr_FirePath = str((vsL_Duration - 300) * (vss_Speed / 60) * 2) & "m"
        Else
            vsStr_FirePath = str(vsL_Duration * (vss_Speed / 60)) & "m"
        End If
    End If
    If Me.OB_ByDuration = True Then
        If ffSng_PointChange(Me.TB_Duration) > 10 Then
            vsStr_FirePath = str((ffSng_PointChange(Me.TB_Duration) - 5) * vss_Speed * 2) & "m"
        Else
            vsStr_FirePath = str(ffSng_PointChange(Me.TB_Duration) * vss_Speed) & "m"
        End If
        vsD_TimeCur = DateAdd("n", ffSng_PointChange(Me.TB_Duration), VmD_TimeStart)
        vsD_TimeCur = DateAdd("s", Ost(ffSng_PointChange(Me.TB_Duration)) * 60, vsD_TimeCur)
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = str(ffSng_PointChange(Me.TB_Radius) * 2) & "m"
        If ffSng_PointChange(Me.TB_Radius) / (vss_Speed * 0.5) > 10 Then
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / vss_Speed + 5) * 60      ' В СЕКУНДАХ!!!!!!
        Else
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / (vss_Speed / 2)) * 60    ' В СЕКУНДАХ!!!!!!
        End If

        vsD_TimeCur = DateAdd("s", vsL_Duration, VmD_TimeStart)
        
    End If
    '---Задаем вычисленные размеры и указываем вычисленную дату
    If vsO_DropMaster.Name = "Сектор 90" Then
        vsStr_FirePath = str(ffSng_PointChange(val(vsStr_FirePath) / 2)) & "m"
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    ElseIf vsO_DropMaster.Name = "Сектор 180" Then
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsStr_FirePath = str(ffSng_PointChange(val(vsStr_FirePath) / 2)) & "m"
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    Else
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    End If
    vsO_DropShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(vsD_TimeCur)) & ")"
    
'---Передаем вброшенной фигуре значение объекта пожара из формы добавления площади
    If Me.OB_SpeedByObject.value = True Then
        vsO_DropShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_DropShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_DropShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
'        vsO_DropShape.Cells("Prop.FireSpeedLine").FormulaU = CDbl(Me.TB_Speed2.Value)
    End If
       
    
'---Очищаем объекты
    Set vsO_DropTargetShape = Nothing
    Set vsO_DropShape = Nothing
    Set vsO_DropMaster = Nothing
    
Exit Sub

EX:
'MsgBox "Одно из указанных вами значений слишком велико или введено с ошибками! " & _
'    "Проверьте правильность введенных данных!", vbCritical
    MsgBox "В процессе работы программы возникла ошибка! Убедитесь в правильности введенных вами данных."
'---Очищаем объекты
    Set vsO_DropTargetShape = Nothing
    Set vsO_DropShape = Nothing
    Set vsO_DropMaster = Nothing
    SaveLog Err, "s_FireShapeDrop"
End Sub


Private Sub s_PrognoseFire()
'Процедура строит прогноз развития пожара, обращает его в зону горения и предает полученной фигуре рассчитанные параметры
Dim vO_Fire As c_Fire
Dim x As Double, y As Double
Dim shp As Visio.Shape
Dim vsO_FireShape As Visio.Shape
Dim vss_Speed As Single 'Скорость распространения огня
Dim vsStr_FirePath As String 'путь пройденный огнем
Dim vsD_TimeCur As Date
Dim vsL_Duration As Long

On Error GoTo Tail

'---Создаем экземпляр класса c_Fire для прогнозирования площади горения
    Set vO_Fire = New c_Fire

'---Определяем фигуру очага пожара и стартовые координаты
    Set shp = Application.ActiveWindow.Selection(1)
    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)

'---Определяем рассчетную скорость распространения огня
    vss_Speed = GetSpeed
'    If Me.OB_SpeedByDirect = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed2))
'    ElseIf Me.OB_SpeedByObject = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed1))
'    End If

'---Определяем размеры фигуры в соответствии с путем пройденным огнем в соответствии _
    с указанными пользователем значениями
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        vsL_Duration = DateDiff("s", VmD_TimeStart, vsD_TimeCur) ' В СЕКУНДАХ!!!!!!
        If vsL_Duration / 60 > 10 Then
            vsStr_FirePath = (vsL_Duration - 300) * (vss_Speed / 60)
        Else
            vsStr_FirePath = (vsL_Duration * (vss_Speed / 60)) / 2
        End If
    End If
    If Me.OB_ByDuration = True Then
        If ffSng_PointChange(Me.TB_Duration) > 10 Then
            vsStr_FirePath = (ffSng_PointChange(Me.TB_Duration) - 5) * vss_Speed
        Else
            vsStr_FirePath = (ffSng_PointChange(Me.TB_Duration) * vss_Speed) / 2
        End If
        vsD_TimeCur = DateAdd("n", ffSng_PointChange(Me.TB_Duration), VmD_TimeStart)
        vsD_TimeCur = DateAdd("s", Ost(ffSng_PointChange(Me.TB_Duration)) * 60, vsD_TimeCur)
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = ffSng_PointChange(Me.TB_Radius)
        If ffSng_PointChange(Me.TB_Radius) / (vss_Speed * 0.5) > 10 Then
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / vss_Speed + 5) * 60      ' В СЕКУНДАХ!!!!!!
        Else
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / (vss_Speed / 2)) * 60    ' В СЕКУНДАХ!!!!!!
        End If

        vsD_TimeCur = DateAdd("s", vsL_Duration, VmD_TimeStart)
        
    End If

'---Непосредственно строим прогноз развития пожара
    With vO_Fire
        .PB_CheckOpens = Me.CB_CheckOpens.value
        .PI_DoorsREI = CInt(ffSng_PointChange(Me.TB_REI))
        .PS_LineSpeedM = vss_Speed
        .S_SetFullShape x, y, vsStr_FirePath
    End With

'---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
    Set vsO_FireShape = Application.ActiveWindow.Selection(1)
    ImportAreaInformation
    

    
'---Передаем вброшенной фигуре значение объекта пожара из формы добавления площади
    If Me.OB_SpeedByObject.value = True Then
        vsO_FireShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_FireShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_FireShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
    End If
    '---Добавляем вычисленную дату/время для полученной фигуры площади горения
    vsO_FireShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(vsD_TimeCur)) & ")"


Set vsO_FireShape = Nothing
Set vO_Fire = Nothing
Set shp = Nothing
Set vO_Fire = Nothing

Exit Sub
Tail:
'    Debug.Print Err.Description
    Set vsO_FireShape = Nothing
    Set vO_Fire = Nothing
    Set shp = Nothing
    Set vO_Fire = Nothing
    SaveLog Err, "s_PrognoseFire"
End Sub













'-------------------------------------Блок служебных/инструментальных процедур и функций-------------------------
Private Function ffSng_PointChange(afStr_String) As Single
'Функция преобразования точек в запятые
Dim i As Integer
Dim vfStr_TempString As String

For i = 1 To Len(afStr_String)
    If Mid(afStr_String, i, 1) = "." Then
        vfStr_TempString = vfStr_TempString & ","
    Else
        vfStr_TempString = vfStr_TempString & Mid(afStr_String, i, 1)
    End If
Next i

ffSng_PointChange = CSng(vfStr_TempString)


End Function

Function Ost(Count As Single) As Single
'Функция возвращает дробную долю числа, с точностью до сотых
Ost = Round(Count - Int(Count), 2)
End Function

Function GetSpeed() As Single
'---Определяем рассчетную скорость распространения огня
    If Me.OB_SpeedByDirect = True Then
        GetSpeed = CSng(ffSng_PointChange(Me.TB_Speed2))
    ElseIf Me.OB_SpeedByObject = True Then
        GetSpeed = CSng(ffSng_PointChange(Me.TB_Speed1))
    End If
End Function

Function GetIntense() As Single
'---Определяем требуемую интенсивность подачи воды
    If Me.OB_SpeedByObject = True Then
        GetIntense = CSng(ffSng_PointChange(Me.TB_Intense1))
    ElseIf Me.OB_SpeedByDirect = True Then
        GetIntense = CSng(ffSng_PointChange(Me.TB_Intense2))
    End If
End Function

Function IsAcceptableMatrixSize(ByVal maxMatrixSize As Long) As Boolean
Dim xCount As Long
Dim yCount As Long
Dim grain As Integer

    
    On Error GoTo EX
    
    grain = Me.txtGrainSize.value

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
    
    IsAcceptableMatrixSize = xCount * yCount < maxMatrixSize
Exit Function
EX:
    IsAcceptableMatrixSize = False
End Function


'--------------------------------Функции проверки корректности введенных данных----------------------------------
Private Function fC_DataCheck() As Boolean
'Базовая функция возвращает логическое значение корректности данных указанных пользователем в форме
    fC_DataCheck = False

    If Me.OB_ByTime.value = True Then
        If fC_DateCorrCheck = False Then Exit Function
        If fC_DateDiffCheck = False Then Exit Function
    End If
    If Me.OB_ByDuration.value = True Then
        If fC_DurationCheck = False Then Exit Function
        If fC_DurationValueCheck = False Then Exit Function
    End If
    If Me.OB_ByRadius.value = True Then
'        If fC_RadiusCheck = False Then Exit Function     !!!1Временно отключено
        If fC_RadiusValueCheck = False Then Exit Function
    End If
    If Me.OB_LiquidShape.value = True Then
        If fB_REICheck = False Then Exit Function
    End If
    
    fC_DataCheck = True

End Function

Private Function fB_REICheck() As Boolean
'Функция возвращает Ложь, если в поле TB_REI указанны некорректные данные
    If IsNumeric(Me.TB_REI.value) Then
        fB_REICheck = True
    Else
        MsgBox "Значение предела огнестойкости, указано некорректно! Проверьте правильность ввода данных", vbCritical
        fB_REICheck = False
    End If

End Function



Private Function fC_DateDiffCheck() As Boolean
'Функция возвращает логическое значение корректности разности времен указанных пользователем в форме и на схеме
Dim vfD_tmpDate As Date

'---Создаем временную переменную
    vfD_tmpDate = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
    If DateDiff("n", VmD_TimeStart, vfD_tmpDate) < 90 And Not DateDiff("n", VmD_TimeStart, vfD_tmpDate) = 0 Then
        fC_DateDiffCheck = True
    ElseIf DateDiff("n", VmD_TimeStart, vfD_tmpDate) = 0 Then
        MsgBox "Разница времени между значением указанным вами и рассчетным, согласно схемы, отсутствует!" & _
        " Это приведет к созданию нулевой площади горения!", vbCritical
        fC_DateDiffCheck = False
    Else
        MsgBox "Разница времени между значением указанным вами и рассчетным, согласно схемы, слишком велика!" & _
        " Это приведет к созданию слишком большой площади горения и может негативно повлиять на производительность системы!", vbCritical
        fC_DateDiffCheck = False
    End If

End Function

Private Function fC_DateCorrCheck() As Boolean
'Функция возвращает логическое значение корректности даты/времени указанных пользователем в форме
Dim vfD_tmpDate As Date
On Error GoTo ErrMsg

'---Пыиаемся создать временную переменную
    vfD_tmpDate = DateValue(Me.TB_Time)
    If vfD_tmpDate > -1 Then 'Если временная переменная создана, то возвращаем истину
        fC_DateCorrCheck = True
    End If
    Exit Function

ErrMsg:
    MsgBox "Указанное вами значение даты/времени не соответствует формату - проверьте правильность вводенных данных!", vbCritical
    fC_DateCorrCheck = False
End Function
Private Function fC_DurationCheck() As Boolean
'Функция возвращает логическое значение корректности правильности указания времени развития указанного пользователем в форме

    If IsNumeric(Me.TB_Duration.value) Then
        fC_DurationCheck = True
    Else
        MsgBox "Значение времени распространения огня, указано некорректно! Проверьте правильность ввода данных", vbCritical
        fC_DurationCheck = False
    End If

End Function

Private Function fC_DurationValueCheck() As Boolean
'Функция возвращает логическое значение корректности количества времени развития указанного пользователем в форме

    If ffSng_PointChange(Me.TB_Duration.value) < 90 And Not ffSng_PointChange(Me.TB_Duration.value) = 0 Then
        fC_DurationValueCheck = True
    ElseIf ffSng_PointChange(Me.TB_Duration.value) = 0 Then
        MsgBox "Указанное вами значение времени распространения огня, равно нулю!" & _
        " Это приведет к созданию нулевой площади горения!", vbCritical
        fC_DurationValueCheck = False
    Else
        MsgBox "Указанное вами значение времени распространения огня, слишком велико!" & _
        " Это приведет к созданию слишком большой площади горения и может негативно повлиять на производительность системы!", vbCritical
        fC_DurationValueCheck = False
    End If

End Function

Private Function fC_RadiusCheck() As Boolean
'Функция возвращает логическое значение корректности радиуса указанного пользователем в форме

    If IsNumeric(Me.TB_Radius.value) Then
        fC_RadiusCheck = True
    Else
        MsgBox "Значение пути пройденного огнем, указано некорректно! Проверьте правильность ввода данных", vbCritical
        fC_RadiusCheck = False
    End If
End Function

Private Function fC_RadiusValueCheck() As Boolean
'Функция возвращает логическое значение корректности размера радиуса указанного пользователем в форме

    If ffSng_PointChange(Me.TB_Radius.value) < 100 And Not ffSng_PointChange(Me.TB_Radius.value) = 0 Then
        fC_RadiusValueCheck = True
    ElseIf ffSng_PointChange(Me.TB_Radius.value) = 0 Then
        MsgBox "Указанное вами значение пути пройденного огнем, равно нулю!" & _
        " Это приведет к созданию нулевой площади горения!", vbCritical
        fC_RadiusValueCheck = False
    Else
        MsgBox "Указанное вами значение пути пройденного огнем, слишком велико!" & _
        " Это приведет к созданию слишком большой площади горения и может негативно повлиять на производительность системы!", vbCritical
        fC_RadiusValueCheck = False
    End If
End Function
