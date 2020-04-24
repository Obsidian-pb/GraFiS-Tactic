Attribute VB_Name = "m_LinesModify"
Option Explicit

'-----------Модуль перестроения графиков в соовтетсвии с внешними данными (анализ или прямое указание из таблицы)
'-----------------------------------------------------------------------------------------------------------------

'--------------------------------График площади горения-----------------------------------------------------------
Public Sub GetFireSquareDataFromAnalize(ByRef shp As Visio.Shape)
'Модификация линии графика ПЛОЩАДИ ТУШЕНИЯ в соответсвии с анализом
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---Получаем от пользователя номер страницы для анализа
    'показываем форму для выбора
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  'Если был нажат кэнсел - выходим не обновляя
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---Получаем массив данных из Анализатора
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---Если в коллекции площадей нет ни одной фигуры (т.е. нет на листе) - выходим из проки
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "На указанной странице нет фигур площадей горения! Получение точных данных невозможно!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---Проверяем, имеется ли фигура очага
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "Фигура очага отсутствует! Получение точных данных невозможно!", vbCritical
        Exit Sub
    End If
    
'---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    
    '---Получаем массив данных
    If shp.Cells("User.IndexPers") = 123 Then
        'Если фигура - график площади пожара
        GraphAnalizer.PS_GetFireSquares DataArray
    ElseIf shp.Cells("User.IndexPers") = 124 Then
        'Если фигура - график площади тушения
        GraphAnalizer.PS_GetExtSquares DataArray
    End If
    
'---Передаем полученный массив процедуре добавления точек графика
    ps_FireGraphicBuild shp, DataArray
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetFireSquareDataFromAnalize"
End Sub

Public Sub GetFireSquareDataFromTable(ByRef shp As Visio.Shape)
'Модификация линии графика ПЛОЩАДИ ГОРЕНИЯ в соответсвии с таблицей данных
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---Если в форме нажат Cancel, то выходим из проки
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    '---Получаем из формы массив данных для построения графика согласно с таблицей данных
    DataForm.PS_GetMainArray MainArray
    '---Перестраиваем график
    ps_FireGraphicBuild shp, MainArray
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetFireSquareDataFromTable"
End Sub


Private Sub ps_FireGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray())
'Прока строит новый график площади горения (тушения)
Dim i As Integer

On Error GoTo EX

    '---Устанавливаем значения для первой (имеющейся) точки
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(MainArray(1, 0)) & "/User.FireMax)*Height"
    
    '---Устанавливаем значения для новых точек (добавляя их)
    For i = 1 To UBound(MainArray, 2)
        PS_FireGraph_AddKnot shp
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(MainArray(1, i)) & "/User.FireMax)*Height"
    Next i
    
Exit Sub
EX:
    MsgBox "Фигуры зоны горения на указанной схеме отсутствуют!", vbCritical
    shp.Delete
End Sub


'--------------------------------График площади пожара-----------------------------------------------------------
Public Sub GetFireTSquareDataFromAnalize(ByRef shp As Visio.Shape)
'Модификация линии графика ПЛОЩАДИ ПОЖАРА в соответсвии с анализом
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---Получаем от пользователя номер страницы для анализа
    'показываем форму для выбора
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  'Если был нажат кэнсел - выходим не обновляя
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---Получаем массив данных из Анализатора
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---Если в коллекции площадей нет ни одной фигуры (т.е. нет на листе) - выходим из проки
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "На указанной странице нет фигур площадей горения! Получение точных данных невозможно!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---Проверяем, имеется ли фигура очага
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "Фигура очага отсутствует! Получение точных данных невозможно!", vbCritical
        Exit Sub
    End If
    
'---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
'        PS_FireGraph_DeleteKnot shp
        PS_FireTGraph_DeleteKnot shp
    Next i
    
    '---Получаем массив данных
    If shp.Cells("User.IndexPers") = 127 Then
        'Если фигура - график площади пожара
        GraphAnalizer.PS_GetFireSquares DataArray
    End If
    
'---Передаем полученный массив процедуре добавления точек графика
    ps_FireTGraphicBuild shp, DataArray
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetFireTSquareDataFromAnalize"
End Sub

Public Sub GetFireTSquareDataFromTable(ByRef shp As Visio.Shape)
'Модификация линии графика ПЛОЩАДИ ТУШЕНИЯ в соответсвии с таблицей данных
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---Если в форме нажат Cancel, то выходим из проки
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    '---Получаем из формы массив данных для построения графика согласно с таблицей данных
    DataForm.PS_GetMainArray MainArray
    '---Перестраиваем график
    ps_FireTGraphicBuild shp, MainArray
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetFireTSquareDataFromTable"
End Sub


Private Sub ps_FireTGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray())
'Прока строит новый график площади пожара (тушения)
Dim i As Integer

On Error GoTo EX

    '---Устанавливаем значения для первой (имеющейся) точки
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(MainArray(1, 0)) & "/User.FireMax)*Height"
    
    '---Устанавливаем значения для новых точек (добавляя их)
    For i = 1 To UBound(MainArray, 2)
        PS_FireTGraph_AddKnot shp
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(MainArray(1, i)) & "/User.FireMax)*Height"
    Next i
    
Exit Sub
EX:
    MsgBox "Фигуры зоны горения на указанной схеме отсутствуют!", vbCritical
    shp.Delete
End Sub



'--------------------------------График расхода-----------------------------------------------------------
Public Sub GetExpenceDataFromAnalize(ByRef shp As Visio.Shape)
'Модификация линии графика РАСХОДА в соответсвии с анализом
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---Получаем от пользователя номер страницы для анализа
    'показываем форму для выбора
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  'Если был нажат кэнсел - выходим не обновляя
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---Получаем массив данных из Анализатора
    '---активируем анализатор
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---Если в коллекции площадей нет ни одной фигуры (т.е. нет на листе) - выходим из проки
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "На указанной странице нет фигур приборов подачи воды! Получение точных данных невозможно!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---Проверяем, имеется ли фигура очага
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "Фигура очага отсутствует! Получение точных данных невозможно!", vbCritical
        Exit Sub
    End If
    
'---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_WaterGraph_DeleteKnot shp
    Next i
    
    '---Получаем массив данных
    If shp.Cells("User.IndexPers") = 125 Then
        'Если фигура - график расхода
        GraphAnalizer.PS_GetWStvolsPodOut DataArray
    ElseIf shp.Cells("User.IndexPers") = 126 Then
        'Если фигура - график эффективного расхода
        GraphAnalizer.PS_GetWStvolsEffPodOut DataArray
    End If
    
'---Передаем полученный массив процедуре добавления точек графика
    ps_ExpenceGraphicBuild shp, DataArray, True

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetExpenceDataFromAnalize"
End Sub

Public Sub GetExpenceDataFromTable(ByRef shp As Visio.Shape)
'Модификация линии графика РАСХОДОВ в соответсвии с таблицей данных
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---Если в форме нажат Cancel, то выходим из проки
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---Оцищаем все точки графика кроме первой (чтоб не исчез вообще)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_WaterGraph_DeleteKnot shp
    Next i

    '---Получаем из формы массив данных для построения графика согласно с таблицей данных
    DataForm.PS_GetMainArray MainArray
    '---Перестраиваем график
    ps_ExpenceGraphicBuild shp, MainArray, False
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetExpenceDataFromTable"
End Sub

Private Sub ps_ExpenceGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray(), ByVal SumOpt As Boolean)
'Прока строит новый график РАСХОДА (ЭФФЕКТИВНОГО РАСХОДА)
'SumOpt: ИСТИНА - учет нарастающим итогом, ЛОЖЬ - учет абсолютных значчений расходов
Dim i As Integer
Dim Expence As Double

On Error GoTo EX

    '---Устанавливаем значения для первой (имеющейся) точки
        Expence = MainArray(1, 0)
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(Expence) & "/User.MaxExpense)*Height"
    
    '---Устанавливаем значения для новых точек (добавляя их)
    For i = 1 To UBound(MainArray, 2)
        PS_WaterGraph_AddKnot shp
        If SumOpt = True Then
            Expence = Expence + MainArray(1, i)  'В случае построения при анализе данных
        Else
            Expence = MainArray(1, i)            'В случае построения при задании данных черех таблицу
        End If
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(Expence) & "/User.MaxExpense)*Height"
    Next i
    
Exit Sub
EX:
   'Exit
   MsgBox "Приборы подачи огнетушащих веществ на указанной схеме отсутствуют!", vbCritical
   shp.Delete
End Sub

'--------------------------------Поле графика-----------------------------------------------------------
Public Sub GetCommonDataFromAnalize(ByRef shp As Visio.Shape)
'Модификация общих сведений в графике
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---Активируем анализатор
'---Получаем от пользователя номер страницы для анализа
    'показываем форму для выбора
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  'Если был нажат кэнсел - выходим не обновляя
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---Получаем массив данных из Анализатора
    '---активируем анализатор
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh

'---Проверяем, имеется ли фигура очага
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "Фигура очага отсутствует! Получение точных данных невозможно!", vbCritical
        Exit Sub
    End If

'---Получаем сведения о пожаре
    '!!!Временно отключаем ошибки - потом ОБЯЗАТЕЛЬНО РАЗОБРАТЬСЯ!!!
    On Error Resume Next
    'начало пожара
'    shp.Cells("Prop.TimeBegin").FormulaU = """" & GraphAnalizer.PF_GetBeginDateTime & """"
    shp.Cells("Prop.TimeBegin").FormulaU = "TheDoc!User.FireTime"
    'максимальная площдь
'    shp.Cells("Prop.FireMax").FormulaForceU = "Guard(" & GraphAnalizer.PF_GetMaxSquare(GraphAnalizer.ColP_Fires.Count) & ")"
    shp.Cells("Prop.FireMax").FormulaForceU = "Guard(" & GraphAnalizer.GetMaxGraphSize(GraphAnalizer.ColP_Fires.Count) & ")"
    'максимальное время
    shp.Cells("Prop.TimeMax").FormulaForceU = "Guard(" & GraphAnalizer.PF_GetTimeEnd(5, "s") / 60 & ")"
    'время окончания
    shp.Cells("Prop.TimeEnd").FormulaU = GraphAnalizer.PF_GetTimeEnd(4, "s") / 60
    'интенсивность
    shp.Cells("Prop.WaterIntense").FormulaForceU = "GUARD(" & str(GraphAnalizer.PF_GetIntence(GraphAnalizer.ColP_Fires.Count)) & ")"
    'показываем, что анализ доолжен проводиться по полученной цифре
    shp.Cells("Prop.WaterIntenseType").Formula = "INDEX(1;Prop.WaterIntenseType.Format)"
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetCommonDataFromAnalize"
End Sub

Public Sub ChangeMaxValues(ByRef shp As Visio.Shape)
'Изменение максимальных значений графика
Dim CurShape As Visio.Shape

    On Error GoTo EX

    '---Закрепляем все графики (на листе) относительно их пропорций!!!
    For Each CurShape In Application.ActivePage.Shapes
        If CurShape.CellExists("User.IndexPers", 0) = True And CurShape.CellExists("User.Version", 0) = True Then
            If CurShape.Cells("User.IndexPers") = 123 Or CurShape.Cells("User.IndexPers") = 124 _
                Or CurShape.Cells("User.IndexPers") = 125 Or CurShape.Cells("User.IndexPers") = 126 _
                Or CurShape.Cells("User.IndexPers") = 127 _
                Then   'Если фигура - фигура новых чертежей (в перспективе еще чего-то)
                
                PS_GraphicsFix CurShape
            End If
        End If
    Next CurShape
    
    '---Показываем форму для текущей фигуры графика
    MaxValuesForm.PS_ShowME shp

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ChangeMaxValues"
End Sub










