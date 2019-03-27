VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Mask"
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
'Private col_PlaceShapes As Collection 'Коллекция вброшенных фигур рабочих мест
Private col_WallsShapes As Collection 'Коллекция стен
Private col_DoorsShapes As Collection 'Коллекция дверей
Private col_WindowsShapes As Collection 'Коллекция окон



'--------------------------Процедуры класса-------------------------------------------------------------------------------------
'Private Sub Class_Initialize()
'    Set vO_App = Visio.Application
'End Sub
'
'Private Sub Class_Terminate()
'    Set vO_App = Nothing
'End Sub

'Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
'    col_RoomsShapes.Add Shape
'End Sub



'-----------------------------------------------------Процедуры создания маски----------------------------------------------------

Public Sub s_makeMask()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
    
    '---Очищаем фигуры в слое маска
    ClearLayer "Маска"
    
    '---Формируем коллекции
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
    Set col_WindowsShapes = New Collection
    
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddShape vO_Shape
    Next vO_Shape
    
    '---Делаем маску стен
    '---Проверяем имеются ли фигуры в коллекции col_WallsShapes
    If col_WallsShapes.count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---Удаляем из маски сетн проемы
    If col_WallsShapes.count = 0 And col_DoorsShapes.count = 0 And col_WindowsShapes.count = 0 Then Exit Sub
    s_deleteWindowsAndDoors


Set vO_Shape = Nothing
End Sub


Private Sub s_makeMaskWalls()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim X1 As Double, Y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

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
            vO_TempShape.Cells("LinePattern") = 0
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

Private Sub s_deleteWindowsAndDoors()
'прока удаляет из маски стен ве проемы - окна и двери
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_WindowsAndDoors As Collection
Dim X1 As Double, Y1 As Double
Dim rect As c_Rect
    
    Set col_WindowsAndDoors = New Collection
    
    'Формируем коллекцию прямоугольников для исключения из маски стен
    '---перебираем все фигуры окон
    For Each vO_Shape In col_WindowsShapes
        Set rect = New c_Rect                       'Создаем новую фигуру прямоугольника
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_WindowsAndDoors.Add vO_TempShape
    Next vO_Shape
    '---перебираем все фигуры дверей
    For Each vO_Shape In col_DoorsShapes
        Set rect = New c_Rect                       'Создаем новую фигуру прямоугольника
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_WindowsAndDoors.Add vO_TempShape
    Next vO_Shape
    
    '---Очищаем выделение
    Application.ActiveWindow.DeselectAll
    '---Добавляем в первоначальное выделение маску стен
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---Перебираем все фигуры прямоугольников и добавляем их в выделение
    For Each vO_Shape In col_WindowsAndDoors
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    
    '---Удаляем проемы из маски
    Application.ActiveWindow.Selection.Subtract
    
    '---Запоминаем маску
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)
    
    '---Добавляем маску в слой "Маска стен"
    pO_MaskShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("Маска") & """"
End Sub





'-----------------------------------------Процедуры определения коллекций-------------------------------------------------------
Private Sub ps_AddShape(ByRef aO_Shape As Visio.Shape)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_WallsShapes СТЕНЫ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddShape vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If
'---Проверяем, является ли фигура фигурой ОКНО
    If PFB_isWindow(aO_Shape) Then
        col_WindowsShapes.Add aO_Shape
        Exit Sub
    End If
'---Проверяем, является ли фигура фигурой ДВЕРЬ
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub






