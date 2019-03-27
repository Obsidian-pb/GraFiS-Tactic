VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_WallDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
'----------------------Класс для отрисовщика стен-----------------------------------------
'-----------------------------------------------------------------------------------------
Public WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1
Public WithEvents wallShape As Visio.Shape
Attribute wallShape.VB_VarHelpID = -1
Public WallDrawerToolButton As Office.CommandBarButton  'Рисование стен



'--------------------Базовые проки и функции класса--------------------------------
Public Function Init() As Boolean
    Set visApp = Application
    Set wallShape = GetSelectedWall
    Set WallDrawerToolButton = Application.CommandBars("Конструкции").Controls("Рисование стен")
    
    If wallShape Is Nothing Then
        Init = False
        Term
    Else
        Init = True
    End If
End Function
Public Function Term() As Boolean
    Set visApp = Nothing
    WallDrawerToolButton.State = msoButtonUp
    Term = True

End Function







Private Sub visApp_ShapeAdded(ByVal Shape As IVShape)
'Прока добавления фигуры стены поверх фигуры линии
Dim shp As Visio.Shape
    
    On Error GoTo EX
    
    '---Проверяем, включен ли инструмент Линия
    If Not LineToolIsSelected Then Exit Sub
    '---Проверяем что добавлена именно линия и она прямая
    If Not ShapeIsLine(Shape) Then Exit Sub

    '---Вбрасываем фигуру стены
    Application.EventsEnabled = False
        Set shp = Application.ActivePage.Drop(wallShape, 0, 0)
    Application.EventsEnabled = True

    '---Указываем конечную и начальную точку новой стены
    With shp
        .Cells("BeginX").Formula = Shape.Cells("BeginX").Formula
        .Cells("BeginY").Formula = Shape.Cells("BeginY").Formula
        .Cells("EndX").Formula = Shape.Cells("EndX").Formula
        .Cells("EndY").Formula = Shape.Cells("EndY").Formula
    End With

    '---Delete Line and clear selection
    Shape.Delete
    '---Снимаем выделение
    Application.ActiveWindow.DeselectAll
Exit Sub
EX:
    Application.EventsEnabled = True
    
    On Error Resume Next
    '---Удаляем линию
    'Shape.Delete
    
    '---Снимаем выделение
    Application.ActiveWindow.DeselectAll
End Sub

Private Function LineToolIsSelected() As Boolean
'Возвращает Истина, если инстурмент "Линия" выбран
Dim cbr As Office.CommandBar
Dim btnTool As Office.CommandBarControl
    
    On Error GoTo EX
    
    Set cbr = Application.CommandBars("Drawing")
    Set btnTool = GetCommandBarTool(cbr, 6509)
    LineToolIsSelected = btnTool.State
Exit Function
EX:
    LineToolIsSelected = False
End Function



'----------------------Работа со стеной----------------------------------
Private Function GetSelectedWall() As Visio.Shape
'Функция возвращает фигуру выбранной стены, если выбрана не стена - возвращает nothing
Dim shp As Visio.Shape

    Set shp = Application.ActiveWindow.Selection(1)
    
    If PFB_isWall(shp) Then
        Set GetSelectedWall = shp
    Else
        Set GetSelectedWall = Nothing
    End If
End Function

Private Sub wallShape_BeforeShapeDelete(ByVal Shape As IVShape)
    Term
End Sub


