VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ManagementStvols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс-обертка для окна учета Поданных стволов

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
Private v_ManagementStvols As f_ManagementStvols
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
    Set v_ManagementStvols = Nothing
    Set app = Nothing
End Sub

Public Sub CloseWindow()
    wAddon.Close
End Sub

Private Sub app_WindowTurnedToPage(ByVal Window As IVWindow)
'Реагируем на переход между страницами
    ps_WindowRefresh v_ManagementStvols
End Sub

Private Sub app_CellChanged(ByVal cell As IVCell)
'Реагируем на изменение ячейки
Dim v_ColumnIndex As Integer
Dim vO_LV As ListView
Dim vO_Shape As Visio.Shape

    'Проверяем имеется ли форма f_ManagementStvols как объект и если нет, то дальше ничего не делаем во избежание ошибки
    If v_ManagementStvols Is Nothing Then Exit Sub

    'Проверяем относится ли фигура с измененной ячейкой к активному листу
    If Not cell.Shape.Parent.Name = Application.ActivePage.Name Then Exit Sub
    
    'Оперделяем объектные переменные
    Set vO_Shape = cell.Shape
    Set vO_LV = v_ManagementStvols.ListView1
    
    If pf_IsInList(cell.Shape.Name) Then
        
        Select Case cell.Name
            Case Is = "Prop.Unit"
                vO_LV.ListItems(cell.Shape.Name).Text = pf_CellValue(cell.Shape, cell.Name)
            Case Is = "Prop.StvolType"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(1).Text = pf_GetStvolType(vO_Shape)
            Case Is = "Prop.StvolTypeS"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(1).Text = pf_GetStvolType(vO_Shape)
            Case Is = "Prop.Call"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(2).Text = pf_GetCall(vO_Shape)
            Case Is = "Prop.SetTime"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(3).Text = Format(vO_Shape.Cells("Prop.SetTime").Result(visNumber), "hh:mm:ss")
            Case Is = "Prop.Personnel"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(4).Text = pf_CellValue(vO_Shape, "Prop.Personnel")
            Case Is = "Prop.UseDirection"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(5).Text = vO_Shape.Cells("Prop.UseDirection").ResultStr(visString)
            Case Is = "Prop.PodOut"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(6).Text = pf_GetPodOut(vO_Shape)
            Case Is = "Prop.PodOutS"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(6).Text = pf_GetPodOut(vO_Shape)
            Case Is = "Prop.TTHType"
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(1).Text = pf_GetStvolType(vO_Shape)
                vO_LV.ListItems(vO_Shape.Name).ListSubItems(6).Text = pf_GetPodOut(vO_Shape)
            Case Is = "Prop.MainManeure"
                'Удаляем из списка
                vO_LV.ListItems.Remove (vO_Shape.Name)  '.ListItems(Cell.Shape.Name).
                
        End Select
    
    Else
    
        '---Если фигура является маневренным стволом, и при этом ее нет в списке (т.е. она была маневернной), то...
        If vO_Shape.CellExists("User.IndexPers", 0) = True And vO_Shape.CellExists("User.Version", 0) = True Then
            If pf_IsStvols(vO_Shape.Cells("User.IndexPers")) And Not pf_IsManeure(vO_Shape) Then
                'Добавляем в список
                pf_AddNewRow vO_LV.ListItems.Count + 1, vO_Shape
            End If
        End If
    
    End If
    
    
End Sub

Private Sub app_ShapeAdded(ByVal Shape As IVShape)
'Учитываем добавление новых фигур
Dim Row As Integer

    On Error GoTo EX

    Row = v_ManagementStvols.ListView1.ListItems.Count + 1

        If Shape.CellExists("User.IndexPers", 0) = True And Shape.CellExists("User.Version", 0) = True Then
            If pf_IsStvols(Shape.Cells("User.IndexPers")) And Not pf_IsManeure(Shape) Then
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

    v_ManagementStvols.ListView1.ListItems.Remove (Shape.Name)
    
Exit Sub
EX:
    'Ошибка
End Sub



'---------------------------------Процедуры работы с формой в окне-------------------------------------------------------
Private Sub ShowWindow()
'Процедура показывает окно "Управление силами и средствами" и ассоциирует его с формой Base
Dim otherForm As Object

    If AnyFormsShowed(otherForm) Then
        Set wAddon = ActiveWindow.Windows.Add("Stvols", visWSVisible + visWSDockedBottom + visWSAnchorMerged, visAnchorBarAddon, , , , , "{BD8994FC-A032-412B-B557-9039E82AB7FA}")
    Else
        Set wAddon = ActiveWindow.Windows.Add("Stvols", visWSVisible + visWSDockedBottom + visWSAnchorMerged, visAnchorBarAddon, , , , 210, "{BD8994FC-A032-412B-B557-9039E82AB7FA}")
    End If
    
    Set v_ManagementStvols = f_ManagementStvols
    v_ManagementStvols.Caption = "Stvols"


    FormHandle = FindWindow(vbNullString, "Stvols")

    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE

    SetParent FormHandle, wAddon.WindowHandle32

'---Настройка ListView
    v_ManagementStvols.ListView1.View = lvwReport
    v_ManagementStvols.ListView1.FullRowSelect = True
    v_ManagementStvols.ListView1.GridLines = True
    v_ManagementStvols.ListView1.LabelEdit = lvwManual
    
    v_ManagementStvols.ListView1.ColumnHeaders.Add 1, "Prop.Unit", "Подразделение", 50
    v_ManagementStvols.ListView1.ColumnHeaders.Add 2, "Prop.StvolType", "Тип ствола", 100
    v_ManagementStvols.ListView1.ColumnHeaders.Add 3, "Prop.Call", "Позывной", 80
    v_ManagementStvols.ListView1.ColumnHeaders.Add 4, "Prop.SetTime", "Время подачи", 90
    v_ManagementStvols.ListView1.ColumnHeaders.Add 5, "Prop.Personnel", "Личный состав", 70
    v_ManagementStvols.ListView1.ColumnHeaders.Add 6, "Prop.UseDirection", "Направление работы", 80
    v_ManagementStvols.ListView1.ColumnHeaders.Add 7, "Prop.PodOut", "Производительность", 80

    wAddon.Caption = "Поданные стволы"
    
'---передача в процедуру получения информации
    ps_GetStvols v_ManagementStvols
    
'---Устанавливаем размер формы, чтобы она выглядела корректно
    If Not otherForm Is Nothing Then
        v_ManagementStvols.Width = otherForm.Width
        v_ManagementStvols.Height = otherForm.Height
    End If
End Sub

Private Function AnyFormsShowed(ByRef frm As Object) As Boolean
'Возвращает True если хоть одна форма уже показана
    If f_ManagementGDZS.Visible = True Then
        AnyFormsShowed = True
        Set frm = f_ManagementGDZS
        Exit Function
    End If
    If f_ManagementTechincs.Visible = True Then
        AnyFormsShowed = True
        Set frm = f_ManagementTechincs
        Exit Function
    End If
AnyFormsShowed = False
End Function

Public Sub PS_ShowWindow()
    If v_ManagementStvols Is Nothing Then
        ShowWindow
    End If
End Sub


Private Sub ps_GetStvols(ByRef StvolForm As f_ManagementStvols)
'Процедура перебирает все фигуры на лисет и в случае, если они удовлетсворяют условиям поиска вставляют и данные в таблицу
Dim vO_Shape As Visio.Shape
Dim Row As Integer

    Row = 1
    
    On Error GoTo EX
    
    For Each vO_Shape In ActivePage.Shapes
        If vO_Shape.CellExists("User.IndexPers", 0) = True And vO_Shape.CellExists("User.Version", 0) = True Then
            If pf_IsStvols(vO_Shape.Cells("User.IndexPers")) And Not pf_IsManeure(vO_Shape) Then
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
    Set v_ManagementStvols = Nothing
End Sub

Private Sub ps_WindowRefresh(ByRef TechForm As f_ManagementStvols)
'Процедура обновления содержимого ListView
On Error GoTo EX
    TechForm.ListView1.Sorted = False
    TechForm.ListView1.ListItems.Clear
    ps_GetStvols TechForm
Exit Sub
EX:
    'Exit
End Sub

Private Sub ps_ChangeLVCellData(ByRef a_Cell As Visio.cell, ByVal a_ColumnIndex As Integer)
'Процедура устанавливает значение измененной ячейки в форме
    v_ManagementStvols.ListView1.ListItems(a_Cell.Shape.Name).ListSubItems(a_ColumnIndex).Text = _
        pf_CellValue(a_Cell.Shape, a_Cell.Name)
End Sub


'-----------------------Функции проверки -----------------------------------------------
Private Function pf_IsStvols(ByVal a_IndexPers As Integer) As Boolean
'Если фигура является стволом - возвращается Истина, в противном случае - Ложь
    If a_IndexPers >= 34 And a_IndexPers <= 39 Or a_IndexPers = 45 Or a_IndexPers = 76 Or a_IndexPers = 77 Then
        pf_IsStvols = True
    Else
        pf_IsStvols = False
    End If
End Function

Private Function pf_IsManeure(ByRef a_Shape As Visio.Shape) As Boolean
'Если фигура имеет свойство - "Маневр" и это свойство равно Истина, то возвращается Истина, в противном случае - Ложь
    pf_IsManeure = False
    If a_Shape.CellExists("Actions.MainManeure.Checked", 0) = True Then
        If a_Shape.Cells("Actions.MainManeure.Checked") = 1 Then pf_IsManeure = True
    End If
    '---Обрабатываем фигуры старых версий
    If a_Shape.CellExists("Prop.MainManeure", 0) = True Then
        If a_Shape.Cells("Prop.MainManeure").ResultStr(visString) = "Маневр" Then pf_IsManeure = True
    End If
End Function

Private Function pf_IsInList(ByVal aS_ShapeName As String) As Boolean
'Функция возвращает Истина, если указанная фигура имеется в списке ListView и Ложь, если нет
Dim i As Integer

pf_IsInList = False
    For i = 1 To v_ManagementStvols.ListView1.ListItems.Count
        If v_ManagementStvols.ListView1.ListItems(i).Key = aS_ShapeName Then pf_IsInList = True
    Next i
End Function

Private Function pf_IsValidCellName(ByVal aS_CellName As String) As Boolean
'Функция возвращает Истина, если указанная ячейка имеется в списке ListView и Ложь, если нет
Dim i As Integer

    pf_IsValidCellName = False
    
    For i = 1 To v_ManagementStvols.ListView1.ColumnHeaders.Count
        If v_ManagementStvols.ListView1.ColumnHeaders(i).Key = aS_CellName Then pf_IsValidCellName = True
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

Private Function pf_GetStvolType(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает название ствола в зависимости от того
On Error GoTo EX

    If aO_Shape.Cells("Prop.TTHType").ResultStr(visString) = "Стандартные" Then
        pf_GetStvolType = aO_Shape.Cells("Prop.StvolTypeS").ResultStr(visString)
    Else
        pf_GetStvolType = aO_Shape.Cells("Prop.StvolType").ResultStr(visString)
    End If
Exit Function
EX:
    pf_GetStvolType = "ошибка"
End Function

Private Function pf_GetPodOut(ByRef aO_Shape As Visio.Shape) As String
'Функция возвращает название ствола в зависимости от того
On Error GoTo EX

    Select Case aO_Shape.Cells("User.IndexPers")
        Case Is = 34  'Водяной ручной ствол
            If aO_Shape.Cells("Prop.TTHType").ResultStr(visString) = "Стандартные" Then
                pf_GetPodOut = aO_Shape.Cells("Prop.PodOutS").Result(visUnitsNone)   'ResultStr(visString)
            Else
                pf_GetPodOut = aO_Shape.Cells("Prop.PodOut").Result(visUnitsNone)
            End If
        Case Is = 36 'Водяной лафетный ствол
            If aO_Shape.Cells("Prop.TTHType").ResultStr(visString) = "Стандартные" Then
                pf_GetPodOut = aO_Shape.Cells("Prop.PodOutS").Result(visUnitsNone)   'ResultStr(visString)
            Else
                pf_GetPodOut = aO_Shape.Cells("Prop.PodOut").Result(visUnitsNone)
            End If
        Case Is = 35 'Пенный ствол
            If aO_Shape.Cells("Prop.TTHType").ResultStr(visString) = "Стандартные" Then
                pf_GetPodOut = aO_Shape.Cells("Prop.FoamOut").Result(visUnitsNone)   'ResultStr(visString)
            Else
                pf_GetPodOut = aO_Shape.Cells("Prop.FoamOut").Result(visUnitsNone)
            End If
        Case Is = 37 'Пенный лафетный ствол
            If aO_Shape.Cells("Prop.TTHType").ResultStr(visString) = "Стандартные" Then
                pf_GetPodOut = aO_Shape.Cells("Prop.FoamOut").Result(visUnitsNone)   'ResultStr(visString)
            Else
                pf_GetPodOut = aO_Shape.Cells("Prop.FoamOut").Result(visUnitsNone)
            End If
    End Select
Exit Function
EX:
    pf_GetPodOut = "ошибка"
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

    Set vO_LV = v_ManagementStvols.ListView1
    
    vO_LV.ListItems.Add a_Row, a_Shape.Name, a_Shape.Cells("Prop.Unit").ResultStr(visString)  'Иконку добавить можно здесь!
    
    vO_LV.ListItems(a_Row).ListSubItems.Add 1, "Prop.StvolType", pf_GetStvolType(a_Shape)
    vO_LV.ListItems(a_Row).ListSubItems.Add 2, "Prop.Call", pf_GetCall(a_Shape)
    vO_LV.ListItems(a_Row).ListSubItems.Add 3, "Prop.SetTime", Format(a_Shape.Cells("Prop.SetTime").Result(visNumber), "hh:mm:ss")
    vO_LV.ListItems(a_Row).ListSubItems.Add 4, "Prop.Personnel", pf_CellValue(a_Shape, "Prop.Personnel")
    vO_LV.ListItems(a_Row).ListSubItems.Add 5, "Prop.UseDirection", a_Shape.Cells("Prop.UseDirection").ResultStr(visString)
    vO_LV.ListItems(a_Row).ListSubItems.Add 6, "Prop.PodOut", pf_GetPodOut(a_Shape)
    
End Sub

