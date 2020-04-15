Attribute VB_Name = "m_Connect"
Private Function GetTechShapeForGESystem(ByRef shp As Visio.Shape, ByRef previousShp As Visio.Shape) As Visio.Shape
'
Dim con As Connect
Dim sideShp As Visio.Shape

    For Each con In shp.Connects
        If Not con.ToSheet = previousShp Then
'            Debug.Print con.ToSheet.Name
'            Set GetTechShapeForGESystem = GetTechShapeForGESystem(con.ToSheet, shp)
            Set sideShp = con.ToSheet
            If sideShp.CellExists("User.IndexPers", 0) Then
                If sideShp.Cells("User.IndexPers") = 40 Then
                    Set GetTechShapeForGESystem = sideShp
'                    Debug.Print sideShp.Name
                    Exit Function
                End If
            End If
            Set GetTechShapeForGESystem = GetTechShapeForGESystem(sideShp, shp)
        End If
    Next con
    For Each con In shp.FromConnects
        If Not con.FromSheet = previousShp Then
'            Debug.Print con.FromSheet.Name
'            Set GetTechShapeForGESystem = GetTechShapeForGESystem(con.FromSheet, shp)
            Set sideShp = con.FromSheet
            If sideShp.CellExists("User.IndexPers", 0) Then
                If sideShp.Cells("User.IndexPers") = 40 Then
                    Set GetTechShapeForGESystem = sideShp
'                    Debug.Print sideShp.Name
                    Exit Function
                End If
            End If
            Set GetTechShapeForGESystem = GetTechShapeForGESystem(sideShp, shp)
        End If
    Next con

'Set GetTechShapeForGESystem = sideShp
End Function

Public Sub GESystemTest(ShpObj As Visio.Shape)

On Error GoTo ex
    If IsWorkWithGidroelevator(ShpObj) Then
        ShpObj.Cells("User.GESystemCheck").FormulaU = 1
    Else
        ShpObj.Cells("User.GESystemCheck").FormulaU = 0
    End If
    
Exit Sub
ex:

End Sub

Public Sub SSS()
'    Debug.Print GetTechShapeForGESystem(Application.ActiveWindow.Selection(1), Application.ActiveWindow.Selection(1))
'    SetWorkWithGidroelevatroOption GetTechShapeForGESystem(Application.ActiveWindow.Selection(1), Application.ActiveWindow.Selection(1)), 1
    Debug.Print IsWorkWithGidroelevator(Application.ActiveWindow.Selection(1))
End Sub

Public Function IsWorkWithGidroelevator(ByRef shp As Visio.Shape) As Boolean
'‘ункци€ возвращает True если это гидроэлеваторна€ система, False, если нет.
    On Error GoTo ex
    
    Debug.Print GetTechShapeForGESystem(shp, shp)
    IsWorkWithGidroelevator = True
    
Exit Function
ex:
    IsWorkWithGidroelevator = False
End Function

Private Sub SetWorkWithGidroelevatroOption(ByRef shp As Visio.Shape, ByVal value As Integer)
'ѕытаемс€ присвоить €чейке "User.WorkWithGE" фигуры значение, если ошибка - ничего не делаем
    On Error GoTo ex
    shp.Cells("User.WorkWithGE").FormulaU = value
ex:
End Sub
