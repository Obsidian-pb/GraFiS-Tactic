VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MaskBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------Класс создает и возвращает фигуру маски слоя-------------------------------
Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Public pO_MaskShape As Visio.Shape                 'фигура маски - суммируется из теней стен и используется как общая тень стен для обрезки
Public pO_FreespaceMaskShape As Visio.Shape        'фигура маски свободного пространства - получается вырезванием из прямоугольника страницы маски стен
Const mmInInch As Double = 25.4


'--------------------------Коллекции--------------------------------------------------------------------------------------------
Private col_WallsShapes As Collection 'Коллекция стен
Private col_DoorsShapes As Collection 'Коллекция дверей
Private col_WindowsShapes As Collection 'Коллекция окон







'-----------------------------------------------------Процедуры создания маски----------------------------------------------------
Public Sub s_makeMask()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
    
    '---Очищаем фигуры в слое маска
    ClearLayer "Маска"
    
    '---Формируем коллекции
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
    
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddShape vO_Shape
    Next vO_Shape
    
    '---Делаем маску стен
    '---Проверяем имеются ли фигуры в коллекции col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---Удаляем из маски стен проемы
    If col_WallsShapes.Count = 0 And col_DoorsShapes.Count = 0 Then Exit Sub
    s_deleteDoors


Set vO_Shape = Nothing
End Sub

Public Sub s_makeOpenSpaceMask()
'Процедура создает маску открытого пространства
Dim pageShape As Visio.Shape
Dim pO_MaskShapeClone As Visio.Shape
Dim pageWidth As Double
Dim pageHeight As Double
    
    '---Определяем размер страницы
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)

    '---Рисуем прямоугольник размером со страницу
    Set pageShape = Application.ActivePage.DrawRectangle(0, 0, pageWidth / mmInInch, pageHeight / mmInInch)
    
    '---Добавляем клон маски стен
    Dim x1 As Double, y1 As Double
    '---Определяем координаты для вброса
    pO_MaskShape.XYToPage pO_MaskShape.Cells("LocPinX").Result(visInches), pO_MaskShape.Cells("LocPinY").Result(visInches), x1, y1
    '---Вбрасываем новую фигуру по координатам
    Set pO_MaskShapeClone = Application.ActivePage.Drop(pO_MaskShape, x1, y1)
       
    '---Добавляем в выделение прямоугольник страницы и маску стен
    Application.ActiveWindow.Selection.DeselectAll
    Application.ActiveWindow.Select pageShape, visSelect
    Application.ActiveWindow.Select pO_MaskShapeClone, visSelect
    '---Вычитаем из прямоугольника листа клон маски стен
    Application.ActiveWindow.Selection.Subtract

    '---Запоминаем маску
    Set pO_FreespaceMaskShape = Application.ActiveWindow.Selection(1)
    
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillForegnd).FormulaU = "THEMEGUARD(TINT(THEME(""FillColor""),24))"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillBkgnd).FormulaU = "THEMEGUARD(SHADE(FillForegnd,LUMDIFF(THEME(""FillColor""),THEME(""FillColor2""))))"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowLine, visLineColorTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillForegndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillShdwForegndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillShdwBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowImage, visImageTransparency).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionCharacter, 0, visCharacterColorTrans).FormulaU = "90%"

End Sub

Private Sub s_makeMaskWalls()
'Процедура создает маску стен
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---перебираем все фигуры стен
    For Each vO_Shape In col_WallsShapes
        '---Определяем координаты для вброса
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
        '---Вбрасываем новую фигуру по координатам
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
            vO_TempShape.Cells("PinX") = x1
            vO_TempShape.Cells("PinY") = y1
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
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
'    SaveLog Err, "s_makeMaskWalls"
    Set col_Shadows = Nothing
End Sub

Private Sub s_deleteDoors()
'прока удаляет из маски стен  двери
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Doors As Collection
Dim x1 As Double, y1 As Double
Dim rect As c_Rect
    
    Set col_Doors = New Collection
    
    'Формируем коллекцию прямоугольников для исключения из маски стен
    '---перебираем все фигуры дверей
    For Each vO_Shape In col_DoorsShapes
        Set rect = New c_Rect                       'Создаем новую фигуру прямоугольника
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_Doors.Add vO_TempShape
    Next vO_Shape
    
    '---Очищаем выделение
    Application.ActiveWindow.DeselectAll
    '---Добавляем в первоначальное выделение маску стен
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---Перебираем все фигуры прямоугольников и добавляем их в выделение
    For Each vO_Shape In col_Doors
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
    If aO_Shape.Shapes.Count > 1 Then
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
''---Проверяем, является ли фигура фигурой ОКНО
'    If PFB_isWindow(aO_Shape) Then
'        col_WindowsShapes.Add aO_Shape
'        Exit Sub
'    End If
'---Проверяем, является ли фигура фигурой ДВЕРЬ
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub


