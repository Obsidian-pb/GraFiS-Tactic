Attribute VB_Name = "m_WorkWithConnections"
Option Explicit
'--------------------------------------------------------------Модуль хранящий процедуры соединения пожарной техники и пожарных рукавов--------------------------------------
Private cpO_InShape As Visio.Shape, cpO_OutShape As Visio.Shape 'Фигуры в которую входит поток и выходит, соответственно


'---Постоянные индеков
Const ccs_InIdent = "Connections.GFS_In"
Const ccs_OutIdent = "Connections.GFS_Ou"
Const vb_ShapeType_Other = 0
Const vb_ShapeType_Hose = 1
Const vb_ShapeType_PTV = 2
Const vb_ShapeType_Razv = 3
Const vb_ShapeType_Tech = 4



Public Sub ConnectionsRefresh(ShpObj As Visio.Shape)
'Процедура освежает подключения рукавов к текущей фигуре
Dim conn As Visio.Connect
Dim i As Integer

On Error Resume Next  ' Не забыть - ошибки пропускаются - для того, чтобы не стопорить программу на старых фигурах!!!

'---Очищаем значения для подключенного лафетного ствола
    If ShpObj.CellExists("User.GFS_OutLafet", 0) = True Then
        ShpObj.Cells("User.GFS_OutLafet").FormulaU = 0
        ShpObj.Cells("User.GFS_OutLafet.Prompt").FormulaU = 0
    End If
    
'---Инициируем i
    i = 0

'---Очищаем значения в ячейках сеrциb Scratch - аннулируем информацию о предыдущих подключениях
    Do While ShpObj.CellsSRCExists(visSectionScratch, i, 0, 0) = True
        If Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_Ou" Then
                ShpObj.CellsSRC(visSectionScratch, i, 4).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 5).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
        ElseIf Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_In" Then
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
        End If
        i = i + 1
        If i > 100 Then Exit Do
    Loop

'---Обновляем соединения рукавов
    For Each conn In ShpObj.FromConnects
        Ps_ConnectionAdd conn
    Next conn

Set conn = Nothing
End Sub


'----------------------------------------Процедуры работы с соединениями-----------------------------------------------
Public Sub Ps_ConnectionAdd(ByRef aO_Conn As Visio.Connect)
'Процедура непосредственно осуществляет соединение фигур
Dim vO_FromShape As Visio.Shape, vO_ToShape As Visio.Shape
Dim vi_InRowNumber As Integer, vi_OutRowNumber As Integer

On Error GoTo EndSub
    
'---Определяем каие фигуры были соединены
    Set vO_FromShape = aO_Conn.FromSheet
    Set vO_ToShape = aO_Conn.ToSheet

'---Проверяем, являются ли соединенные фигуры фигурами ГраФиС
    If vO_FromShape.CellExists("User.IndexPers", 0) = False Or _
        vO_ToShape.CellExists("User.IndexPers", 0) = False Then Exit Sub '---Проверяем являются ли фигуры _
                                                                                фигурами ГаФиС
'---Проверяем, являются ли соединенные фигуры элементами НРС
    If f_IdentShape(vO_FromShape.Cells("User.IndexPers").Result(visNumber)) = 0 Or _
        f_IdentShape(vO_ToShape.Cells("User.IndexPers").Result(visNumber)) = 0 Then Exit Sub

'---Идентифицируем подающую и принимающую фигуры - при соединении рукавов и ПТВ!!!
    '---Для From фигуры
    If Left(aO_Conn.FromCell.Name, 18) = ccs_InIdent Then
        Set cpO_InShape = aO_Conn.FromSheet
        Set cpO_OutShape = aO_Conn.ToSheet
        vi_InRowNumber = aO_Conn.FromCell.row
        vi_OutRowNumber = aO_Conn.ToCell.row
    ElseIf Left(aO_Conn.FromCell.Name, 18) = ccs_OutIdent Then
        Set cpO_InShape = aO_Conn.ToSheet
        Set cpO_OutShape = aO_Conn.FromSheet
        vi_InRowNumber = aO_Conn.ToCell.row
        vi_OutRowNumber = aO_Conn.FromCell.row
    End If
    '---Для То фигуры
    If Left(aO_Conn.ToCell.Name, 18) = ccs_InIdent Then
        Set cpO_InShape = aO_Conn.ToSheet
        Set cpO_OutShape = aO_Conn.FromSheet
        vi_InRowNumber = aO_Conn.ToCell.row
        vi_OutRowNumber = aO_Conn.FromCell.row
    ElseIf Left(aO_Conn.ToCell.Name, 18) = ccs_OutIdent Then
        Set cpO_InShape = aO_Conn.FromSheet
        Set cpO_OutShape = aO_Conn.ToSheet
        vi_InRowNumber = aO_Conn.FromCell.row
        vi_OutRowNumber = aO_Conn.ToCell.row
    End If
    '---В случае, если обе фигуры - рукава
    If vO_FromShape.Cells("User.IndexPers") = 100 And _
        vO_ToShape.Cells("User.IndexPers") = 100 Then
        '---Проверяем у какой фигура рукава входящий поток больше
        If aO_Conn.ToSheet.Cells("Scratch.D1") > aO_Conn.FromSheet.Cells("Scratch.D1") Then
            Set cpO_InShape = aO_Conn.ToSheet
            Set cpO_OutShape = aO_Conn.FromSheet
            vi_InRowNumber = aO_Conn.ToCell.row
            vi_OutRowNumber = aO_Conn.FromCell.row
        Else
            '---Проверяем у какой фигуры рукава ID выше (вброшена позже)
            If aO_Conn.ToSheet.ID > aO_Conn.FromSheet.ID Then
                Set cpO_InShape = aO_Conn.ToSheet
                Set cpO_OutShape = aO_Conn.FromSheet
                vi_InRowNumber = aO_Conn.ToCell.row
                vi_OutRowNumber = aO_Conn.FromCell.row
            Else
                Set cpO_InShape = aO_Conn.FromSheet
                Set cpO_OutShape = aO_Conn.ToSheet
                vi_InRowNumber = aO_Conn.FromCell.row
                vi_OutRowNumber = aO_Conn.ToCell.row
            End If
        End If
    End If

    '---Запускаем процедуру связывания данных в фигурах
       ps_LinkShapes vi_InRowNumber, vi_OutRowNumber

Exit Sub

EndSub:
    Resume Next
'    Debug.Print Err.Description
    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
    SaveLog Err, "Ps_ConnectionAdd"
End Sub

Private Sub ps_LinkShapes(ByVal ai_InRowNumber As Integer, ByVal ai_OutRowNumber As Integer)
'Внутренняя процедура - связывает данные в соединяемых фигурах
Dim vi_IPInShape As Integer, vi_IPOutShape As Integer
Dim vb_InShapeType As Byte, vb_OutShapeType As Byte
Dim i As Integer
Dim vs_Formula As String

On Error GoTo ExitSub

'---Проверить чем являются соединяемые фигуры
    '---УстанавливаемIndexPers для каждой из фигур
    vi_IPInShape = cpO_InShape.Cells("User.IndexPers")
    vi_IPOutShape = cpO_OutShape.Cells("User.IndexPers")
    '---Проверяем
        '---Для принимающей фигуры
        vb_InShapeType = f_IdentShape(vi_IPInShape)
        '---Для подающей фигуры
        vb_OutShapeType = f_IdentShape(vi_IPOutShape)
        
'---В зависимости от типа соединяемых фигур выбрать процедуру связывания
    '---Рукав->ПТВ
        If vb_OutShapeType = vb_ShapeType_Hose And (vb_InShapeType = vb_ShapeType_PTV Or vb_InShapeType = vb_ShapeType_Razv) Then
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Дополнение для лафетных водяных стволов:
            If vi_IPInShape = 36 Or vi_IPInShape = 37 Or vi_IPInShape = 39 Then
            '---Указываем, что к фигуре подсоединены рукава на вход
            cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 1
            End If
            '---Привязываем маневренность линии к маневренности ствола
            If cpO_InShape.CellExists("Actions.MainManeure", 0) = True Then
                cpO_OutShape.Cells("Prop.ManeverHose").Formula = "INDEX(Sheet." & cpO_InShape.ID _
                    & "!Actions.MainManeure.Checked" & ";Prop.ManeverHose.Format)"
            End If
        End If
    '---ПТВ->Рукав
        If (vb_OutShapeType = vb_ShapeType_PTV Or vb_OutShapeType = vb_ShapeType_Razv) And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
        End If
    '---Рукав->Рукав
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.C1"
            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.LineTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            cpO_OutShape.Cells("Prop.ManeverHose").FormulaU = "Sheet." & cpO_InShape.ID & "!Prop.ManeverHose"
        End If
    '---Рукав->Техника основная
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Tech Then
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.ArrivalTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Указываем, что к фигуре подсоединены рукава на вход
            If cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 1
            ElseIf cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 2
            End If
        End If
    '---Техника основная->Рукав
        If vb_OutShapeType = vb_ShapeType_Tech And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            '---Указываем, что к фигуре подсоединены рукава на выход
            If cpO_OutShape.Cells("Scratch.A" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_OutShape.Cells("Scratch.A" & CStr(ai_OutRowNumber + 1)).Formula = 1
            ElseIf cpO_OutShape.Cells("Scratch.A" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_OutShape.Cells("Scratch.A" & CStr(ai_OutRowNumber + 1)).Formula = 2
            End If
        End If
    
    
    
    
Exit Sub
ExitSub:
'    Debug.Print Err.Description
    SaveLog Err, "ps_LinkShapes"
End Sub





'----------------------------------------Служебные Функции-----------------------------------------------
Private Function f_IdentShape(ByVal ai_ShapeIP As Integer) As Integer
'Функция идентифициурет фигуру и возвращает значение её типа
Dim Arr_PTVs(12, 2)
Dim i As Integer

'---Указываем значения IndexPers и соответствующие им определения
    Arr_PTVs(0, 0) = 34  'Водяной ручной ствол
        Arr_PTVs(0, 1) = vb_ShapeType_PTV
    Arr_PTVs(1, 0) = 35  'Пенный ствол
        Arr_PTVs(1, 1) = vb_ShapeType_PTV
    Arr_PTVs(2, 0) = 36  'Лафетный водяной
        Arr_PTVs(2, 1) = vb_ShapeType_PTV
    Arr_PTVs(3, 0) = 37  'Лафетный пенный
        Arr_PTVs(3, 1) = vb_ShapeType_PTV
    Arr_PTVs(4, 0) = 39  'Возимый лафетный ствол
        Arr_PTVs(4, 1) = vb_ShapeType_PTV
    Arr_PTVs(5, 0) = 42  'Разветвление
        Arr_PTVs(5, 1) = vb_ShapeType_Razv
    Arr_PTVs(6, 0) = 45  'Пеноподъемник
        Arr_PTVs(6, 1) = vb_ShapeType_PTV
    Arr_PTVs(7, 0) = 72  'Колонка
        Arr_PTVs(7, 1) = vb_ShapeType_PTV
    Arr_PTVs(8, 0) = 100 'Напорная линия
        Arr_PTVs(8, 1) = vb_ShapeType_Hose
    Arr_PTVs(9, 0) = 1 'Автоцистерна пожарная
        Arr_PTVs(9, 1) = vb_ShapeType_Tech
    Arr_PTVs(10, 0) = 2 'АНР
        Arr_PTVs(10, 1) = vb_ShapeType_Tech
    Arr_PTVs(11, 0) = 20 'Рукавный автомобиль
        Arr_PTVs(11, 1) = vb_ShapeType_Tech
        
'---Указываем значение по умолчанию
    f_IdentShape = vb_ShapeType_Other

'---Проверяем является ли фигура
        For i = 0 To 11
            If ai_ShapeIP = Arr_PTVs(i, 0) Then f_IdentShape = Arr_PTVs(i, 1): Exit Function
        Next i

End Function
