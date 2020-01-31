Attribute VB_Name = "archiv"
'Sub CloneSecMiscellanious(ShapeFromID As Integer, ShapeToID As Integer)
''Процедура копирования значений строк для секции "Miscellanious"
''---Объявляем переменные
'Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
'Dim RowCountFrom As Integer, RowCountTo As Integer
'Dim RowNum As Integer
'
''---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
'Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)
'Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)
'
''---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
'    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
'        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
'    Next j
'
'End Sub



'------------------------------------------Соединение фигур основа-----------------------------
'Dim vO_FromShape, vO_ToShape As Visio.Shape
'Dim vi_InRowNumber, vi_OutRowNumber As Integer
'
''---Определяем каие фигуры были соединены
'    Set vO_FromShape = Connects.FromSheet
'    Set vO_ToShape = Connects.ToSheet
'
''---Проверяем, являются ли соединенные фигуры фигурами ГраФиС
'    If vO_FromShape.CellExists("User.IndexPers", 0) = False Or _
'        vO_ToShape.CellExists("User.IndexPers", 0) = False Then Exit Sub '---Проверяем являются ли фигуры _
'                                                                                фигурами ГаФиС
''---Проверяем, являются ли соединенные фигуры элементами НРС
'    If f_IdentShape(vO_FromShape.Cells("User.IndexPers").Result(visNumber)) = 0 Or _
'        f_IdentShape(vO_ToShape.Cells("User.IndexPers").Result(visNumber)) = 0 Then Exit Sub
'
''---Идентифицируем подающую и принимающую фигуры - при соединении рукавов и ПТВ!!!
'    '---Для From фигуры
'    If Left(Connects(1).FromCell.Name, 18) = ccs_InIdent Then
'        Set cpO_InShape = Connects.FromSheet
'        Set cpO_OutShape = Connects.ToSheet
'        vi_InRowNumber = Connects(1).FromCell.Row
'        vi_OutRowNumber = Connects(1).ToCell.Row
'    ElseIf Left(Connects(1).FromCell.Name, 18) = ccs_OutIdent Then
'        Set cpO_InShape = Connects.ToSheet
'        Set cpO_OutShape = Connects.FromSheet
'        vi_InRowNumber = Connects(1).ToCell.Row
'        vi_OutRowNumber = Connects(1).FromCell.Row
'    End If
'    '---Для То фигуры
'    If Left(Connects(1).ToCell.Name, 18) = ccs_InIdent Then
'        Set cpO_InShape = Connects.ToSheet
'        Set cpO_OutShape = Connects.FromSheet
'        vi_InRowNumber = Connects(1).ToCell.Row
'        vi_OutRowNumber = Connects(1).FromCell.Row
'    ElseIf Left(Connects(1).ToCell.Name, 18) = ccs_OutIdent Then
'        Set cpO_InShape = Connects.FromSheet
'        Set cpO_OutShape = Connects.ToSheet
'        vi_InRowNumber = Connects(1).FromCell.Row
'        vi_OutRowNumber = Connects(1).ToCell.Row
'    End If
'
'    '---Запускаем процедуру связывания данных в фигурах
'       ps_LinkShapes vi_InRowNumber, vi_OutRowNumber
'
'
'    On Error Resume Next
''    Debug.Print "Принимающая фигура: " & cpO_InShape.Name
''    Debug.Print "Подающая фигура: " & cpO_OutShape.Name
''    Debug.Print "Фигура рукава: " & cpO_HoseShape.Name
''    Debug.Print Left(Connects(1).FromCell.Name, 18) & " -> " & Left(Connects(1).ToCell.Name, 18)
''    Debug.Print vO_FromShape & " -> " & vO_ToShape
'    Set cpO_InShape = Nothing
'    Set cpO_OutShape = Nothing


'----------------Для подключения к лафетным стволам
''                Debug.Print cpO_InShape.Cells("User.Connects")
'                cpO_InShape.Cells("User.Connects").Formula = cpO_InShape.Cells("User.Connects") + 1
'                '---Проверяем количество подключенных рукавов
'                If cpO_InShape.Cells("User.Connects") > 1 Then
'                    cpO_InShape.Cells("Scratch.D1").Formula = "User.PodOut/2"
'                    cpO_InShape.Cells("Scratch.D2").Formula = "User.PodOut/2"
'                Else
'                    cpO_InShape.Cells("Scratch.D" & CStr(ai_InRowNumber + 1)).Formula = "User.PodOut"
'                End If
