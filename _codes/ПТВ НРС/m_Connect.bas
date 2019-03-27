Attribute VB_Name = "m_Connect"
Public Sub ColConn(ShpObj As Visio.Shape)
'Процедура привязки получаемого потока !Колонки! от фигуры к которой она приклеена (если ПГ)
Dim ToShape As Integer

'---Предотвращаем появление сообщения об ошибке
On Error GoTo SubExit

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
            ShpObj.Cells("Prop.FlowIn").FormulaU = "Sheet." & ToShape & "!Prop.Production" & ""
    Else
        ShpObj.Cells("Prop.FlowIn").FormulaU = 0
    End If

Exit Sub

SubExit:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ColConn"
End Sub


'Private Function GetTechShapeForGESystem(ByRef shp As Visio.Shape, ByRef previousShp As Visio.Shape) As Visio.Shape
''
'Dim con As Connect
'Dim sideShp As Visio.Shape
'
'    For Each con In shp.Connects
'        If Not con.ToSheet = previousShp Then
''            Debug.Print con.ToSheet.Name
''            Set GetTechShapeForGESystem = GetTechShapeForGESystem(con.ToSheet, shp)
'            Set sideShp = con.ToSheet
'            If sideShp.CellExists("User.IndexPers", 0) Then
'                If sideShp.Cells("User.IndexPers") = 40 Then
'                    Set GetTechShapeForGESystem = sideShp
''                    Debug.Print sideShp.Name
'                    Exit Function
'                End If
'            End If
'            Set GetTechShapeForGESystem = GetTechShapeForGESystem(sideShp, shp)
'        End If
'    Next con
'    For Each con In shp.FromConnects
'        If Not con.FromSheet = previousShp Then
''            Debug.Print con.FromSheet.Name
''            Set GetTechShapeForGESystem = GetTechShapeForGESystem(con.FromSheet, shp)
'            Set sideShp = con.FromSheet
'            If sideShp.CellExists("User.IndexPers", 0) Then
'                If sideShp.Cells("User.IndexPers") = 40 Then
'                    Set GetTechShapeForGESystem = sideShp
''                    Debug.Print sideShp.Name
'                    Exit Function
'                End If
'            End If
'            Set GetTechShapeForGESystem = GetTechShapeForGESystem(sideShp, shp)
'        End If
'    Next con
'
''Set GetTechShapeForGESystem = sideShp
'End Function
'
'Public Sub GESystemTest(ShpObj As Visio.Shape)
'
'    If IsWorkWithGidroelevator(ShpObj) Then
'        Debug.Print "true"
'    End If
'
'End Sub
'
'Public Sub SSS()
''    Debug.Print GetTechShapeForGESystem(Application.ActiveWindow.Selection(1), Application.ActiveWindow.Selection(1))
''    SetWorkWithGidroelevatroOption GetTechShapeForGESystem(Application.ActiveWindow.Selection(1), Application.ActiveWindow.Selection(1)), 1
'    Debug.Print IsWorkWithGidroelevator(Application.ActiveWindow.Selection(1))
'End Sub
'
'Public Function IsWorkWithGidroelevator(ByRef shp As Visio.Shape) As Boolean
''Функция возвращает True если это гидроэлеваторная система, False, если нет.
'    On Error GoTo EX
'
'    Debug.Print GetTechShapeForGESystem(shp, shp)
'    IsWorkWithGidroelevator = True
'
'Exit Function
'EX:
'    IsWorkWithGidroelevator = False
'End Function
'
'Private Sub SetWorkWithGidroelevatroOption(ByRef shp As Visio.Shape, ByVal value As Integer)
''Пытаемся присвоить ячейке "User.WorkWithGE" фигуры значение, если ошибка - ничего не делаем
'    On Error GoTo EX
'    shp.Cells("User.WorkWithGE").FormulaU = value
'EX:
'End Sub

'------------------------!!!АРХИВ!!!-------------------------------------------
'Private Function GetTechShapeForGESystem(ByRef shp As Visio.Shape, ByRef previousShp As Visio.Shape) As Visio.Shape
''
'Dim con As Connect
'Dim sideShp As Visio.Shape
'
'    For Each con In shp.Connects
''        Debug.Print con.FromSheet.Name
'
'        If Not con.ToSheet = previousShp Then
'            Debug.Print con.ToSheet.Name
'            GetTechShapeForGESystem con.ToSheet, shp
'        End If
'    Next con
'    For Each con In shp.FromConnects
''        Debug.Print con.FromSheet.Name
'
'        If Not con.FromSheet = previousShp Then
'            Debug.Print con.FromSheet.Name
'            GetTechShapeForGESystem con.FromSheet, shp
'        End If
'    Next con
'
'
'End Function
'
'Public Sub SSS()
'    GetTechShapeForGESystem Application.ActiveWindow.Selection(1), Application.ActiveWindow.Selection(1)
'End Sub
'------------------------!!!АРХИВ!!!-------------------------------------------
