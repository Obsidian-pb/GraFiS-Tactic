VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ManagementTimeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс-обертка для окна учета ГДЗС

'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

#If VBA7 Then
    Public FormHandle As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                    ByVal lpClassName As String, _
                    ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIndex As LongPtr, _
                    ByVal dwNewLong As Long) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetParent Lib "user32" ( _
                    ByVal hWndChild As LongPtr, _
                    ByVal hWndNewParent As LongPtr) As LongPtr
#Else
    Public FormHandle As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                    ByVal lpClassName As String, _
                    ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hwnd As Long, _
                    ByVal nIndex As Long) As Long
    Private Declare Function SetParent Lib "user32" ( _
                    ByVal hWndChild As Long, _
                    ByVal hWndNewParent As Long) As Long
#End If

Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private v_ManagementTimeLine As f_ManagementTimeLine
Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Private Const con_BorderWidth = 6
Private Const con_ToolBarHeight = 50





'-------------------------------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    ShowWindow
    Set app = Visio.Application
End Sub

Private Sub Class_Terminate()
    Set v_ManagementTimeLine = Nothing
    Set app = Nothing
End Sub

Private Sub app_WindowTurnedToPage(ByVal Window As IVWindow)
'Реагируем на переход между страницами
    ps_WindowRefresh v_ManagementTimeLine
End Sub

Private Sub app_CellChanged(ByVal cell As IVCell)
'Реагируем на изменение ячейки
Dim v_ColumnIndex As Integer
Dim vO_LV As ListView
Dim vO_Shape As Visio.Shape

    'Проверяем имеется ли форма v_ManagementTimeLine как объект и если нет, то дальше ничего не делаем во избежание ошибки
    If v_ManagementTimeLine Is Nothing Then Exit Sub

    'Проверяем относится ли фигура с измененной ячейкой к активному листу
    If Not cell.Shape.Parent.Name = Application.ActivePage.Name Then Exit Sub

    'Оперделяем объектные переменные
    Set vO_Shape = cell.Shape
    Set vO_LV = v_ManagementTimeLine.ListView1

    If pf_IsInList(cell.Shape.Name) Then

        Select Case cell.Name
            Case Is = "Prop.Unit"
                vO_LV.ListItems(vO_Shape.Name).Text = pf_CellValue(vO_Shape, cell.Name)
            Case Is = "Prop.Call"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(1).Text = pf_GetCall(vO_Shape)
            Case Is = "User.IndexPers.Prompt"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(2).Text = pf_GetTimeLineType(vO_Shape)
            Case Is = "Prop.ArrivalTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.ArrivalTime").ResultStr(visDate)
            Case Is = "Prop.LineTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.LineTime").ResultStr(visDate)
            Case Is = "Prop.SetTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.SetTime").ResultStr(visDate)
            Case Is = "Prop.FormingTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.FormingTime").ResultStr(visDate)
            Case Is = "Prop.SquareTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.SquareTime").ResultStr(visDate)
            Case Is = "Prop.FireTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.FireTime").ResultStr(visDate)
            
            Case Is = "Actions.MainManeure.Checked"
                'Удаляем из списка
                vO_LV.ListItems.Remove (vO_Shape.Name)  '.ListItems(Cell.Shape.Name).

        End Select

    End If


End Sub


'---------------------------------Процедуры работы с формой в окне-------------------------------------------------------
Private Sub ShowWindow()
'Процедура показывает окно "Управление силами и средствами" и ассоциирует его с формой Base
'Dim FormHandle As Long

'    Set wAddon = ActiveWindow.Windows.Add("TimeLine", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)
    Set wAddon = ActiveWindow.Windows.Add("TimeLine", visWSVisible + visWSDockedRight, visAnchorBarAddon, , , 450, 300)
    
    Set v_ManagementTimeLine = New f_ManagementTimeLine
    v_ManagementTimeLine.Caption = "TimeLine"


    FormHandle = FindWindow(vbNullString, "TimeLine")

    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE

    SetParent FormHandle, wAddon.WindowHandle32

'---Настройка ListView
    v_ManagementTimeLine.ListView1.View = lvwReport
    v_ManagementTimeLine.ListView1.FullRowSelect = True
    v_ManagementTimeLine.ListView1.GridLines = True
    v_ManagementTimeLine.ListView1.LabelEdit = lvwManual
    
    v_ManagementTimeLine.ListView1.ColumnHeaders.Add 1, "Prop.Unit", "Подразделение", 75
    v_ManagementTimeLine.ListView1.ColumnHeaders.Add 2, "Prop.Call", "Позывной", 70
    v_ManagementTimeLine.ListView1.ColumnHeaders.Add 3, "Type", "Тип", 70
    v_ManagementTimeLine.ListView1.ColumnHeaders.Add 4, "Prop.FormingTime", "Время", 90

    wAddon.Caption = "Обстановка по времени"
    
'---передача в процедуру получения информации
    ps_GetTimeLine v_ManagementTimeLine

'---Сортируем по времени
    v_ManagementTimeLine.ps_SortBy (5)

End Sub

Private Sub app_ShapeAdded(ByVal Shape As IVShape)
'Учитываем добавление новых фигур
Dim Row As Integer

    On Error GoTo EX

    Row = v_ManagementTimeLine.ListView1.ListItems.Count + 1

        If Shape.CellExists("User.IndexPers", 0) = True And Shape.CellExists("User.Version", 0) = True Then
            If pf_IsTimeLine(Shape) And Not pf_IsManeure(Shape) Then
                pf_AddNewRow Row, Shape
            End If
        End If

Exit Sub
EX:
    'Ошибка
End Sub

Private Sub app_BeforeShapeDelete(ByVal Shape As IVShape)
'Учитываем удаление фигуры

    On Error GoTo EX

    v_ManagementTimeLine.ListView1.ListItems.Remove (Shape.Name)
    
Exit Sub
EX:
    'Ошибка
End Sub

Public Sub PS_ShowWindow()
    If v_ManagementTimeLine Is Nothing Then
        ShowWindow
    End If
End Sub


Private Sub ps_GetTimeLine(ByRef StvolForm As f_ManagementTimeLine)
'Процедура перебирает все фигуры на лисет и в случае, если они удовлетсворяют условиям поиска вставляют и данные в таблицу
Dim vO_Shape As Visio.Shape
Dim Row As Integer

    Row = 1
    
'On Error GoTo EX
    
    For Each vO_Shape In ActivePage.Shapes
        If vO_Shape.CellExists("User.IndexPers", 0) = True And vO_Shape.CellExists("User.Version", 0) = True Then
            If pf_IsTimeLine(vO_Shape) And Not pf_IsManeure(vO_Shape) Then
                pf_AddNewRow Row, vO_Shape
                
                Row = Row + 1
            End If
        End If
    Next

Exit Sub
EX:
    'Ошибка
End Sub

Private Sub wAddon_BeforeWindowClosed(ByVal Window As IVWindow)
    Set v_ManagementTimeLine = Nothing
End Sub

Private Sub ps_WindowRefresh(ByRef TechForm As f_ManagementTimeLine)
'Процедура обновления содержимого ListView
On Error GoTo EX
    TechForm.ListView1.Sorted = False
    TechForm.ListView1.ListItems.Clear
    ps_GetTimeLine TechForm
Exit Sub
EX:
    'Exit
End Sub



'-----------------------Функции проверки -----------------------------------------------
Private Function pf_IsTimeLine(ByRef a_Shape As Visio.Shape) As Boolean
'Если фигура является фигурой таймлайна - возвращается Истина, в противном случае - Ложь
    If a_Shape.CellExists("Prop.ArrivalTime", 0) = True Or a_Shape.CellExists("Prop.LineTime", 0) = True _
        Or a_Shape.CellExists("Prop.SetTime", 0) = True Or a_Shape.CellExists("Prop.FormingTime", 0) = True _
        Or a_Shape.CellExists("Prop.SquareTime", 0) = True Or a_Shape.CellExists("Prop.FireTime", 0) = True _
        Then
        pf_IsTimeLine = True
    Else
        pf_IsTimeLine = False
    End If
End Function

Private Function pf_IsManeure(ByRef a_Shape As Visio.Shape) As Boolean
'Если фигура имеет свойство - "Маневр" и это свойство равно Истина, то возвращается Истина, в противном случае - Ложь
    pf_IsManeure = False
    If a_Shape.CellExists("Actions.MainManeure.Checked", 0) = True Then
        If a_Shape.Cells("Actions.MainManeure.Checked") = 1 Then pf_IsManeure = True
    End If
End Function

Private Function pf_IsInList(ByVal aS_ShapeName As String) As Boolean
'Функция возвращает Истина, если указанная фигура имеется в списке ListView и Ложь, если нет
Dim i As Integer

pf_IsInList = False
    For i = 1 To v_ManagementTimeLine.ListView1.ListItems.Count
        If v_ManagementTimeLine.ListView1.ListItems(i).Key = aS_ShapeName Then pf_IsInList = True
    Next i
End Function

Private Function pf_IsValidCellName(ByVal aS_CellName As String) As Boolean
'Функция возвращает Истина, если указанная ячейка имеется в списке ListView и Ложь, если нет
Dim i As Integer

    pf_IsValidCellName = False
    
    For i = 1 To v_ManagementTimeLine.ListView1.ColumnHeaders.Count
        If v_ManagementTimeLine.ListView1.ColumnHeaders(i).Key = aS_CellName Then pf_IsValidCellName = True
    Next i

End Function


'-------------------------Функции получения данных----------------------------------------
Private Function pf_CellValue(ByRef aO_Shape As Visio.Shape, ByVal aS_CellName As String) As String
'Функция возвращает значение указанной ячейки указанной фигуры
    pf_CellValue = "-"
    If aO_Shape.CellExists(aS_CellName, 0) = True Then
        pf_CellValue = aO_Shape.Cells(aS_CellName).ResultStr(visString)
        If IsNumeric(pf_CellValue) Then
            pf_CellValue = Int(pf_CellValue)
        End If
    End If
End Function

Private Function pf_GetUnit(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает позывной (если он есть)
On Error GoTo EX

    pf_GetUnit = CallToStr(aO_Shape.Cells("Prop.Unit").ResultStr(visString))
Exit Function
EX:
    pf_GetUnit = "-"
End Function

Private Function pf_GetTimeLineType(ByRef aO_Shape As Visio.Shape) As String
'В зависимости от того, чем является фигура - возвращает такую подпсиь (звено, пост или КПП)
On Error GoTo EX
    pf_GetTimeLineType = aO_Shape.Cells("User.IndexPers.Prompt").ResultStr(visString)
Exit Function
EX:
    pf_GetTimeLineType = "не известно"
End Function


Private Function pf_GetCall(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает позывной (если он есть)
On Error GoTo EX

    pf_GetCall = CallToStr(aO_Shape.Cells("Prop.Call").ResultStr(visString))
Exit Function
EX:
    pf_GetCall = "не доступно"
End Function

Private Function pf_GetTime(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает время фигуры (основное)
On Error GoTo EX

    If aO_Shape.CellExists("Prop.ArrivalTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.ArrivalTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.LineTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.LineTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SetTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SetTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FormingTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FormingTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SquareTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SquareTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FireTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FireTime").ResultStr(visDate)
        Exit Function
    End If

Exit Function
EX:
    pf_GetTime = "не доступно"
End Function


Private Sub pf_AddNewRow(ByVal a_Row As Integer, ByRef a_Shape As Visio.Shape)
'Процедура добавляет новую строку согласно данных фигуры
Dim vO_LV As ListView

    Set vO_LV = v_ManagementTimeLine.ListView1
    
    vO_LV.ListItems.Add a_Row, a_Shape.Name, pf_GetUnit(a_Shape)  'Иконку добавить можно здесь!
    
    vO_LV.ListItems(a_Row).ListSubItems.Add 1, "Prop.Call", pf_GetCall(a_Shape)
    vO_LV.ListItems(a_Row).ListSubItems.Add 2, "Prop.Type", pf_GetTimeLineType(a_Shape)
    vO_LV.ListItems(a_Row).ListSubItems.Add 3, "Prop.Time", pf_GetTime(a_Shape)

End Sub



