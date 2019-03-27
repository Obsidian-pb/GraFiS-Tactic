VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_WorkPlaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Private pO_MaskShape As Visio.Shape 'фигура маски - суммируется из теней стен и используется как общая тень стен для обрезки

'--------------------------Коллекции--------------------------------------------------------------------------------------------
Private col_RoomsShapes As Collection 'Коллекция вброшенных фигур комнат
Private col_PlaceShapes As Collection 'Коллекция вброшенных фигур рабочих мест
Private col_WallsShapes As Collection 'Коллекция стен




'--------------------------Процедуры класса-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set vO_App = Visio.Application
End Sub

Private Sub Class_Terminate()
    Set vO_App = Nothing
End Sub

Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
    col_RoomsShapes.Add Shape
End Sub


'-----------------------------------Основные процедуры------------------------------------------------------------------------

Public Sub S_SetFullShape()
'Основная процедура запуска вброса перечня рабочих мест на листе
Dim vO_SheetShape As Visio.Shape


On Error GoTo Tail

    '---Определяем коллекции фигур
        Set col_RoomsShapes = New Collection
        Set col_PlaceShapes = New Collection
        Set col_WallsShapes = New Collection
    
    '---ОТКЛЮЧАЕМ реагирование объекта приложения на события (вброс новой фигуры)
        vO_App.EventsEnabled = False
    
    '---Формируем маску
        s_makeMask
    
    '---Вбрасываем чистую фигуру с размерами листа
        Set vO_SheetShape = Application.ActiveWindow.Page.DrawRectangle(0, 0, _
                        Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches), _
                        Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches))
    
    '---Вырезаем из фигуры листа кмонаты
        Application.ActiveWindow.DeselectAll
            Application.ActiveWindow.Select vO_SheetShape, visSelect
            Application.ActiveWindow.Select pO_MaskShape, visSelect
        Application.ActiveWindow.Selection.Subtract
        '---ВКЛЮЧАЕМ реагирование объекта приложения на события (вброс новой фигуры)
        vO_App.EventsEnabled = True
        '---Фрагментируем вброшенную фигуру
        Application.ActiveWindow.Selection.Fragment
    
    '---Отбрасываем обрезки фигур (внешний контур и фигуры стен - у всех них больше одной графической секции)
        ps_DeleteTrash
       
    '---Заменяем полученные комнаты фигурами мест
        '---ОТКЛЮЧАЕМ реагирование объекта приложения на события (вброс новой фигуры)
        vO_App.EventsEnabled = False
        'Вставляем фигуры рабочих мест
        ps_InsertWorkPlaces
        'Удаляем все фигуры комнат
        ps_DeleteRooms
    
    '---ВКЛЮЧАЕМ реагирование объекта приложения на события (вброс новой фигуры)
    vO_App.EventsEnabled = True
    Set col_RoomsShapes = Nothing
    Set col_PlaceShapes = Nothing
    Set col_WallsShapes = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "S_SetFullShape"
    '---ВКЛЮЧАЕМ реагирование объекта приложения на события (вброс новой фигуры)
    vO_App.EventsEnabled = True
    Set col_RoomsShapes = Nothing
    Set col_PlaceShapes = Nothing
    Set col_WallsShapes = Nothing
End Sub

Private Sub ps_DeleteTrash()
'Прока удаляет фигуры обрезков
Dim vO_Shape As Visio.Shape
Dim tmpCol As Collection

    Set tmpCol = New Collection
    
    'Удаляем обрезки и одновременно формируем временную коллекцию комнат
    For Each vO_Shape In col_RoomsShapes
        If PFI_FirstSectionCount(vO_Shape) > 1 Then
            vO_Shape.Delete
        Else
            tmpCol.Add vO_Shape
        End If
    Next vO_Shape

    '---Обновляем коллекцию комнат
    Set col_RoomsShapes = tmpCol
End Sub

Private Sub ps_InsertWorkPlaces()
'Прока заменяет все фигуры комнат фигурами рабочих мест
Dim vO_RoomShape As Visio.Shape
Dim vO_PlaceShape As Visio.Shape
Dim vO_Master As Visio.Master
Dim i As Integer
    
'    Set vO_Master = Application.Documents("WALL_M.vss").Masters(4) '("Место")
    Set vO_Master = Application.Documents("WALL_M.vss").Masters("Место")
    
On Error Resume Next
    
    i = 1
    '---Перебираем все фигуры в коллекциикомнат
    For Each vO_RoomShape In col_RoomsShapes
        'вбрасываем фигуру рабочего места
        Set vO_PlaceShape = Application.ActivePage.Drop(vO_Master, 0, 0)
        'задаем ей геометрию комнаты
        PF_GeometryCopy vO_RoomShape, vO_PlaceShape
        'Вычисляем площадь фигуры
        vO_PlaceShape.Cells("User.visFMArea").FormulaU = str(vO_PlaceShape.AreaIU)
        'Перемещаем фигуру на задний план
'        vO_PlaceShape.SendToBack
        'добавляем фигуру места в коллекцию и присваиваем ей очередной порядковый номер
        col_PlaceShapes.Add vO_PlaceShape
        vO_PlaceShape.Cells("Prop.LocationID").FormulaU = i
        
        i = i + 1
    Next vO_RoomShape

End Sub

Private Sub ps_DeleteRooms()
'Прока удаляет все фигуры комнат и очищает коллекцию
Dim vO_RoomShape As Visio.Shape

For Each vO_RoomShape In col_RoomsShapes
    vO_RoomShape.Delete
    col_RoomsShapes.Remove 1
Next vO_RoomShape

End Sub

Public Sub PF_GeometryCopy(ByRef a_RoomShape As Visio.Shape, ByRef a_PlaceShape As Visio.Shape)
'Прока придает фигуре рабочего места геометрию фигуры комнаты
Dim j As Integer
Dim k As Integer
    
    On Error GoTo Tail
    
    '---Формируем геометрию
    '---Очищаем имеющиеся точки
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
    '---Копируем геометрию из комнаты
        j = 1
        Do While a_RoomShape.RowExists(visSectionFirstComponent, j, 0)
            a_PlaceShape.AddRow visSectionFirstComponent, j, a_RoomShape.RowType(visSectionFirstComponent, j)
            
            k = 0
            Do While a_RoomShape.CellsSRCExists(visSectionFirstComponent, j, k, 0)
                a_PlaceShape.CellsSRC(visSectionFirstComponent, j, k).FormulaU = _
                    a_RoomShape.CellsSRC(visSectionFirstComponent, j, k).FormulaU
                
                k = k + 1
            Loop
            j = j + 1
        Loop
    '---Приравниваем положение и размеры исходной фигуры к данным реплики
        a_PlaceShape.Cells("Width").FormulaU = a_RoomShape.Cells("Width").FormulaU
        a_PlaceShape.Cells("Height").FormulaU = a_RoomShape.Cells("Height").FormulaU
        a_PlaceShape.Cells("LocPinX").FormulaU = a_RoomShape.Cells("LocPinX").FormulaU
        a_PlaceShape.Cells("LocPinY").FormulaU = a_RoomShape.Cells("LocPinY").FormulaU
        a_PlaceShape.Cells("PinX").FormulaU = a_RoomShape.Cells("PinX").FormulaU
        a_PlaceShape.Cells("PinY").FormulaU = a_RoomShape.Cells("PinY").FormulaU

Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "PF_GeometryCopy"
End Sub



'-----------------------------------------------------Процедуры создания маски----------------------------------------------------

Private Sub s_makeMask()
'Процедура создает маску обрезки (ОСНОВНАЯ)
Dim vO_Shape As Visio.Shape

    '---Формируем коллекции
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddWall vO_Shape
    Next vO_Shape

    '---Делаем маску стен
   s_makeMaskWalls

Set vO_Shape = Nothing
End Sub



Private Sub s_makeMaskWalls()
'Процедура создает маску обрезки стен
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim X1 As Double, Y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---Проверяем имеются ли фигуры в коллекции col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub

'---перебираем все фигуры стен
    For Each vO_Shape In col_WallsShapes
        '---Определяем координаты для вброса
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), X1, Y1
        '---Вбрасываем новую фигуру по координатам
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), X1, Y1)
            vO_TempShape.Cells("PinX") = X1
            vO_TempShape.Cells("PinY") = Y1
            vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
            
        '---Добавляем во временную коллекцию
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---Формируем на основе собранной коллекции маску
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Union
    Application.ActiveWindow.Selection.SendToBack

'---Запоминаем маску
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

    Set col_Shadows = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_makeMaskWalls"
    Set col_Shadows = Nothing
End Sub


'-----------------------------------------Процедуры определения коллекций-------------------------------------------------------
Private Sub ps_AddWall(ByRef aO_Shape As Visio.Shape)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_WallsShapes СТЕНЫ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddWall vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If

End Sub


'---------------------------------------Служебные функции и проки--------------------------------------------------
Private Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'Функция возвращаетугол относительно родительского элемента
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

Set Shape = Nothing
End Function


