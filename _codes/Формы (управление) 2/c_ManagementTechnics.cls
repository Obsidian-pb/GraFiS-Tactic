VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ManagementTechnics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс-обертка для окна учета Пожарной техники

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
Private v_ManagementTechincs As f_ManagementTechincs
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
    Set v_ManagementTechincs = Nothing
    Set app = Nothing
End Sub

Private Sub app_WindowTurnedToPage(ByVal Window As IVWindow)
'Реагируем на переход между страницами
    ps_WindowRefresh v_ManagementTechincs
End Sub

Private Sub app_CellChanged(ByVal cell As IVCell)
'Реагируем на изменение ячейки
Dim v_ColumnIndex As Integer
Dim vO_LV As ListView
Dim vO_Shape As Visio.Shape

    On Error GoTo EX

    'Проверяем имеется ли форма f_ManagementTechnics как объект и если нет, то дальше ничего не делаем во избежание ошибки
    If v_ManagementTechincs Is Nothing Then Exit Sub

    'Проверяем относится ли фигура с измененной ячейкой к активному листу
    If Not cell.Shape.Parent.Name = Application.ActivePage.Name Then Exit Sub
    
    'Оперделяем объектные переменные
    Set vO_Shape = cell.Shape
    Set vO_LV = v_ManagementTechincs.ListView1
    
    If pf_IsInList(vO_Shape.Name) Then
        
        Select Case cell.Name
            Case Is = "Prop.Unit"
                vO_LV.ListItems(vO_Shape.Name).Text = pf_CellValue(vO_Shape, cell.Name)
            Case Is = "Prop.Call"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(1).Text = pf_GetCall(vO_Shape)
            Case Is = "Prop.Model"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(2).Text = vO_Shape.Cells("Prop.Model").ResultStr(visString)
            Case Is = "Prop.ArrivalTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = vO_Shape.Cells("Prop.ArrivalTime").ResultStr(visString)
            Case Is = "Prop.PersonnelHave"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(4).Text = pf_CellValue(vO_Shape, "Prop.PersonnelHave")
            Case Is = "Actions.MainManeure.Checked"
                'Удаляем из списка
                vO_LV.ListItems.Remove (vO_Shape.Name)
        End Select
    Else
        '---Если фигура является маневренной техникой, и при этом ее нет в списке (т.е. она была маневернной), то...
        If vO_Shape.CellExists("User.IndexPers", 0) = True And vO_Shape.CellExists("User.Version", 0) = True Then
            If pf_IsMainTechnics(vO_Shape.Cells("User.IndexPers")) And Not pf_IsManeure(vO_Shape) Then
                'Добавляем в список
                pf_AddNewRow vO_LV.ListItems.Count + 1, vO_Shape
            End If
        End If
    
    End If

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "app_CellChanged"
End Sub

Private Sub app_ShapeAdded(ByVal Shape As IVShape)
'Учитываем добавление новых фигур
Dim Row As Integer

    On Error GoTo EX
    
    If IsNull(v_ManagementTechincs.ListView1) Then Exit Sub
    
    Row = v_ManagementTechincs.ListView1.ListItems.Count + 1

        If Shape.CellExists("User.IndexPers", 0) = True And Shape.CellExists("User.Version", 0) = True Then
            If pf_IsMainTechnics(Shape.Cells("User.IndexPers")) And Not pf_IsManeure(Shape) Then
                pf_AddNewRow Row, Shape
            End If
        End If

Exit Sub
EX:
    'Ошибка
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
'    SaveLog Err, "app_ShapeAdded"
End Sub

Private Sub app_BeforeShapeDelete(ByVal Shape As IVShape)
'Учитываем удаление фигуры

    On Error GoTo EX

    v_ManagementTechincs.ListView1.ListItems.Remove (Shape.Name)
    
Exit Sub
EX:
    'Ошибка
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
'    SaveLog Err, "app_BeforeShapeDelete"
End Sub


'---------------------------------Процедуры работы с формой в окне-------------------------------------------------------
Private Sub ShowWindow()
'Процедура показывает окно "Управление силами и средствами" и ассоциирует его с формой Base
'Dim FormHandle As LongLong

    On Error GoTo EX

    Set wAddon = ActiveWindow.Windows.Add("Base", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)
    Set v_ManagementTechincs = New f_ManagementTechincs
    v_ManagementTechincs.Caption = "Base"


    FormHandle = FindWindow(vbNullString, "Base")

    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE

    SetParent FormHandle, wAddon.WindowHandle32

'---Настройка ListView
    v_ManagementTechincs.ListView1.View = lvwReport
    v_ManagementTechincs.ListView1.FullRowSelect = True
    v_ManagementTechincs.ListView1.GridLines = True
    v_ManagementTechincs.ListView1.LabelEdit = lvwManual
    
    v_ManagementTechincs.ListView1.ColumnHeaders.Add 1, "Prop.Unit", "Подразделение", 50
    v_ManagementTechincs.ListView1.ColumnHeaders.Add 2, "Prop.Call", "Позывной", 80
    v_ManagementTechincs.ListView1.ColumnHeaders.Add 3, "Prop.Model", "Модель", 100
    v_ManagementTechincs.ListView1.ColumnHeaders.Add 4, "Prop.ArrivalTime", "Время прибытия", 90
    v_ManagementTechincs.ListView1.ColumnHeaders.Add 5, "Prop.PersonnelHave", "Личный состав", 70

    wAddon.Caption = "Прибывшие подразделения"
    
'---передача в процедуру получения информации
    ps_GetTechnics v_ManagementTechincs

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ShowWindow"
End Sub

Public Sub PS_ShowWindow()
    If v_ManagementTechincs Is Nothing Then
        ShowWindow
    End If
End Sub


Private Sub ps_GetTechnics(ByRef TechForm As f_ManagementTechincs)
'Процедура перебирает все фигуры на листе и в случае, если они удовлетсворяют условиям поиска вставляют и данные в таблицу
Dim vO_Shape As Visio.Shape
Dim Row As Integer

    Row = 1
    
On Error GoTo EX
    
    For Each vO_Shape In ActivePage.Shapes
        If vO_Shape.CellExists("User.IndexPers", 0) = True And vO_Shape.CellExists("User.Version", 0) = True Then
            If pf_IsMainTechnics(vO_Shape.Cells("User.IndexPers")) And Not pf_IsManeure(vO_Shape) Then
                pf_AddNewRow Row, vO_Shape
                
                Row = Row + 1
            End If
        End If
    Next

Exit Sub
EX:
    'Ошибка
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ps_GetTechnics"
End Sub

Private Sub wAddon_BeforeWindowClosed(ByVal Window As IVWindow)
    Set v_ManagementTechincs = Nothing
End Sub

Private Sub ps_WindowRefresh(ByRef TechForm As f_ManagementTechincs)
'Процедура обновления содержимого ListView
On Error GoTo EX
    TechForm.ListView1.Sorted = False
    TechForm.ListView1.ListItems.Clear
    ps_GetTechnics TechForm
Exit Sub
EX:
    'Exit
End Sub



'-----------------------Функции проверки -----------------------------------------------
Private Function pf_IsMainTechnics(ByVal a_IndexPers As Integer) As Boolean
'Если фигура является техникой - возвращается Истина, в противном случае - Ложь
    If a_IndexPers <= 20 Or a_IndexPers = 24 Or a_IndexPers = 25 Or a_IndexPers = 26 Or a_IndexPers = 27 Or _
        a_IndexPers = 28 Or a_IndexPers = 29 Or a_IndexPers = 30 Or a_IndexPers = 31 Or a_IndexPers = 32 Or _
        a_IndexPers = 161 Or a_IndexPers = 162 Then
        pf_IsMainTechnics = True
    Else
        pf_IsMainTechnics = False
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
    For i = 1 To v_ManagementTechincs.ListView1.ListItems.Count
        If v_ManagementTechincs.ListView1.ListItems(i).Key = aS_ShapeName Then pf_IsInList = True
    Next i
End Function

Private Function pf_IsValidCellName(ByVal aS_CellName As String) As Boolean
'Функция возвращает Истина, если указанная ячейка имеется в списке ListView и Ложь, если нет
Dim i As Integer

    pf_IsValidCellName = False
    
    For i = 1 To v_ManagementTechincs.ListView1.ColumnHeaders.Count
        If v_ManagementTechincs.ListView1.ColumnHeaders(i).Key = aS_CellName Then pf_IsValidCellName = True
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

Private Function pf_GetCall(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает позывной (если он есть)
On Error GoTo EX

    pf_GetCall = CallToStr(aO_Shape.Cells("Prop.Call").ResultStr(visString))
Exit Function
EX:
    pf_GetCall = "не доступно"
End Function

Private Sub pf_AddNewRow(ByVal a_Row As Integer, ByRef a_Shape As Visio.Shape)
'Процедура добавляет новую строку согласно данных фигуры
Dim vO_LV As ListView

    Set vO_LV = v_ManagementTechincs.ListView1
    
    vO_LV.ListItems.Add a_Row, a_Shape.Name, a_Shape.Cells("Prop.Unit").ResultStr(visString)
    
    vO_LV.ListItems(a_Row).ListSubItems.Add 1, "Prop.Call", CallToStr(a_Shape.Cells("Prop.Call").ResultStr(visString))
    vO_LV.ListItems(a_Row).ListSubItems.Add 2, "Prop.Model", a_Shape.Cells("Prop.Model").ResultStr(visString)
    vO_LV.ListItems(a_Row).ListSubItems.Add 3, "Prop.ArrivalTime", a_Shape.Cells("Prop.ArrivalTime").ResultStr(visString)
    vO_LV.ListItems(a_Row).ListSubItems.Add 4, "Prop.PersonnelHave", pf_CellValue(a_Shape, "Prop.PersonnelHave")

End Sub
