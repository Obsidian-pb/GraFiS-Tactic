VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Integer                           'ID блока - нужен для идентификации, начинается с 1.

Private cellsMatrix As Variant                  'Матрица клеток

Public extinguishingCells As Collection         'Коллекция клеток зоны площади тушения, относящихся к данному блоку
Public fireFrontCells As Collection             'Коллекция клеток фронта пожара, относящихся к даннойму блоку
Public activeCells As Collection                'Коллекция клеток активной зоны (еще не горящих на максимум), относящихся к даннойму блоку

Public fireWay As Single                        'Глубина тушения поданными стволами (в миллиметрах)
Public grain As Long
Public intenseNeed As Single                    'Требуемая интенсивность подачи воды л/(с*м.кв.)
Public expenseFact As Single                    'Объем подаваемой воды л/с

Public alreadyCalculated As Boolean             'Флаг, что блок уже рассчитан (заполнены коллекции)

'--------------------Класс блока зоны тушения (модель)---------------------------------------------------



Private Sub Class_Initialize()
    Set extinguishingCells = New Collection
    Set fireFrontCells = New Collection
    Set activeCells = New Collection
    
    alreadyCalculated = False
End Sub

'-------------------------------Работа с матрицей---------------------------------
Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim cellsMatrix(x, y)
End Sub

Public Sub AddToExtCellsCollection(ByRef coll As Collection)
'Добавляем коллекцию входящих клеток в коллекцию горящих
Dim cell As c_Cell
    
    AddCellsToCollection extinguishingCells, coll
    
    'Для всех входящих клеток устанавливаем значение "в зоне тушения"
    For Each cell In coll
        cellsMatrix(cell.x, cell.y) = ExtCellType.inExtSquare
    Next cell
End Sub

Public Sub AddToActiveCellsCollection(ByRef coll As Collection)
'Добавляем коллекцию входящих клеток в коллекцию активных клеток (внешних, разгорающихся)
Dim cell As c_Cell
    
    AddCellsToCollection activeCells, coll
    
    'Для всех входящих клеток устанавливаем значение "в зоне тушения"
    For Each cell In coll
        cellsMatrix(cell.x, cell.y) = ExtCellType.inExtSquare
    Next cell
End Sub

Public Function IsCellExtinguished(ByVal x As Long, ByVal y As Long) As Boolean
    IsCellExtinguished = cellsMatrix(x, y) = ExtCellType.inExtSquare
End Function



Public Function GetSquare() As Single
'Возвращает площадь зоны тушения в метрах квадратных
    GetSquare = (extinguishingCells.Count * grain ^ 2) / 1000000
End Function

Public Function GetExpenceNeed() As Single
'Возвращает требуемый расход воды для тушения данной площади
    GetExpenceNeed = GetSquare * intenseNeed
End Function

Public Function GetWaterExpenseKind() As WaterExpenseKind
'Подается воды: 0 не подается, 1 - подается недостаточно, 2 - пожается достаточно.
Dim expenseNeed As Single
    
    expenseNeed = GetExpenceNeed
    If expenseFact = 0 Then
        GetWaterExpenseKind = notSet
    End If
    If expenseFact < expenseNeed Then
        GetWaterExpenseKind = notSufficient
    End If
    If expenseFact >= expenseNeed Then
        GetWaterExpenseKind = sufficient
    End If
    
End Function


Public Function IsCellInThisFront(ByRef cell As c_Cell) As Boolean
'Проверяем, относится ли пришедшая клетка к коллекции клеток периметра поажара данного ьлока
Dim cll As c_Cell
    For Each cll In fireFrontCells
        If cll.x = cell.x And cll.y = cell.y Then
            IsCellInThisFront = True
            Exit Function
        End If
    Next
IsCellInThisFront = False
End Function





'--------------------------Свойства---------------------------------------------------
Property Let LetID(ByVal val As Integer)
    ID = val
End Property
Property Get GetID()
    GetID = ID
End Property




'------------------------Инструменты-------------------------------------------------
Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'Добавляем все клетки из коллекции 2 в коллекцию 1
Dim cell As c_Cell

    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub
