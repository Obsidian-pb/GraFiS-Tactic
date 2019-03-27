Attribute VB_Name = "m_WorkWithConnections"
Option Explicit
'--------------------------------------------------------------Модуль хранящий процедуры соединения ПТВ и пожарных рукавов--------------------------------------
Private cpO_InShape As Visio.Shape, cpO_OutShape As Visio.Shape 'Фигуры в которую входит поток и выходит, соответственно


'---Постоянные индеков
Const ccs_InIdent = "Connections.GFS_In"
Const ccs_OutIdent = "Connections.GFS_Ou"
Const vb_ShapeType_Other = 0                'Ничего
Const vb_ShapeType_Hose = 1                 'Рукава
Const vb_ShapeType_PTV = 2                  'ПТВ
Const vb_ShapeType_Razv = 3                 'Разветвление
Const vb_ShapeType_Tech = 4                 'Техника
Const vb_ShapeType_VsasSet = 5              'Всасывающая сетка с линией
Const vb_ShapeType_GE = 6                   'Гидроэлеватор
Const vb_ShapeType_WaterContainer = 7       'Водяная емкость



'----------------------------------------Процедуры работы с соединениями-----------------------------------------------
Public Sub ConnectionsRefresh(ShpObj As Visio.Shape)
'Процедура освежает подключения рукавов к текущей фигуре
Dim Conn As Visio.Connect
Dim i As Integer

On Error Resume Next  ' Не забыть - ошибки пропускаются - для того, чтобы не стопорить программу на старых фигурах!!!

'---Очищаем значения для подключенного лафетного ствола
'    If ShpObj.CellExists("User.GFS_OutLafet", 0) = True Then
'        ShpObj.Cells("User.GFS_OutLafet").FormulaU = 0
'        ShpObj.Cells("User.GFS_OutLafet.Prompt").FormulaU = 0
'    End If
    
'---Инициируем i
    i = 0

'---Очищаем значения в ячейках секции Scratch - аннулируем информацию о предыдущих подключениях
'    Do While ShpObj.CellsSRCExists(visSectionScratch, i, 0, 0) = True  'Для секции Scratch!!!
    Do While ShpObj.CellsSRCExists(visSectionConnectionPts, i, 0, 0) = True  'Для секции Connections!!!
        If Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_Ou" Then
'                ShpObj.CellsSRC(visSectionScratch, i, 4).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 3).FormulaU = 0
        ElseIf Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_In" Then
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
        End If
        ShpObj.CellsSRC(visSectionScratch, i, 0).FormulaU = 0
        i = i + 1
        If i > 100 Then Exit Do
    Loop

'---Обновляем соединения рукавов
    For Each Conn In ShpObj.FromConnects
        Ps_ConnectionAdd Conn
    Next Conn

Set Conn = Nothing
End Sub


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
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "Ps_ConnectionAdd"
    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
End Sub

Private Sub ps_LinkShapes(ByVal ai_InRowNumber As Integer, ByVal ai_OutRowNumber As Integer)
'Внутренняя процедура - связывает данные в соединяемых фигурах
Dim vi_IPInShape, vi_IPOutShape As Integer
Dim vb_InShapeType, vb_OutShapeType As Byte
Dim i As Integer
Dim vs_Formula As String

On Error GoTo EX

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
        If vb_OutShapeType = vb_ShapeType_Hose And (vb_InShapeType = vb_ShapeType_PTV Or vb_InShapeType = vb_ShapeType_Razv Or vb_InShapeType = vb_ShapeType_GE) Then
            'НАПОР ПОЛУЧАЕМ ОТ ВХОДЯЩЕГО ИСТОЧНИКА
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula      'РАСХОД ОСТАЕТСЯ ТАК
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Дополнение для лафетных водяных стволов и пенного ствола:
            If vi_IPInShape = 36 Or vi_IPInShape = 37 Or vi_IPInShape = 39 Or vi_IPInShape = 35 Then
            '---Указываем, что к фигуре подсоединены рукава на вход
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            End If
            '---Привязываем маневренность линии к маневренности ствола
            If cpO_InShape.CellExists("Actions.MainManeure", 0) = True Then
                cpO_OutShape.Cells("Prop.ManeverHose").Formula = "INDEX(Sheet." & cpO_InShape.ID _
                    & "!Actions.MainManeure.Checked" & ";Prop.ManeverHose.Format)"
            End If
            '---Привязываем цвета ПТВ к цветам рукавной линии
            cpO_InShape.Cells("LineColor").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
            cpO_InShape.Cells("FillForegnd").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
            cpO_InShape.Cells("Char.Color").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
        End If
    '---ПТВ->Рукав
        If (vb_OutShapeType = vb_ShapeType_PTV Or vb_OutShapeType = vb_ShapeType_Razv Or vb_OutShapeType = vb_ShapeType_GE) And vb_InShapeType = vb_ShapeType_Hose Then
            
            'ДОБАВЛЯЕМ ПЕРЕДАЧУ НАПОРА РУКАВУ
            cpO_InShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C" & ai_OutRowNumber + 1
            
            'ПЕРЕДАЛАТЬ ЛОГИКУ САМОГО РАЗВЕТВЛЕНИЯ
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
            'ДОБАВЛЯЕМ ПОЛУЧЕНИЕ РАСХОДА ОТ РУКАВА
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, visScratchB).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
        End If
    '---Рукав->Рукав
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Hose Then
            'Было
'            cpO_OutShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            'Стало
            cpO_InShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.LineTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            cpO_OutShape.Cells("Prop.ManeverHose").FormulaU = "Sheet." & cpO_InShape.ID & "!Prop.ManeverHose"
        End If
    '---Рукав->Техника основная
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Tech Then
            'Было
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
'            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            'Стало
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.ArrivalTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Указываем, что к фигуре подсоединены рукава на вход
            If cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            ElseIf cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "2m"
            End If
            '---Проверяем, не получает ли рукав воду от того же МСП, которому отдает
            If SelfWaterGetCheck(cpO_OutShape) Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = 0
            End If
        End If
    '---Техника основная->Рукав
        If vb_OutShapeType = vb_ShapeType_Tech And vb_InShapeType = vb_ShapeType_Hose Then
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 3).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("Scratch.A1").Formula = _
                "Sheet." & cpO_OutShape.ID & "!Scratch.C" & ai_OutRowNumber + 1
                
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            '---Указываем, что к фигуре подсоединены рукава на выход
            If cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "1m"
            ElseIf cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "2m"
            End If
        End If
    '---Всасывающая сетка->Техника основная
        If vb_OutShapeType = vb_ShapeType_VsasSet And vb_InShapeType = vb_ShapeType_Tech Then
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
'            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Указываем, что к фигуре подсоединены рукава на вход
            If cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            ElseIf cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "2m"
            End If
        End If
    '---Рукав->Емкость
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_WaterContainer Then
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).FormulaU = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
        End If
    '---Емкость->Рукав
        If vb_OutShapeType = vb_ShapeType_WaterContainer And vb_InShapeType = vb_ShapeType_Hose Then
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 3).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
'            cpO_InShape.Cells("Scratch.A1").Formula =  Вакууметрический напор?
            
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            '---Указываем, что к фигуре подсоединены рукава на выход
            If cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "1m"
            ElseIf cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "2m"
            End If
        End If
    Exit Sub
    
    
    
    
Exit Sub
EX:
    Debug.Print Err.description
'    Debug.Print vs_Formula
    SaveLog Err, "ps_LinkShapes"

End Sub


'----------------------------------------Служебные процедуры-----------------------------------------------
Private Function SelfWaterGetCheck(ByRef hoseShp As Visio.Shape) As Boolean
'Прока проверяет, не осуществляется ли самозабор и если осуществляется, подключение к всасывающим патрубкам МСП обнуляется
Dim con As Visio.Connect
Dim shpName1 As String
Dim shpName2 As String
    
'---Проверяем какие фигуры подключены к концам рукава
    For Each con In hoseShp.Connects
        If con.FromCell.Name = "BeginX" Then
            shpName1 = con.ToSheet.NameU
        End If
        If con.FromCell.Name = "EndX" Then
            shpName2 = con.ToSheet.NameU
        End If
    Next con
    
'---Если подключена одна и та же фигура - возвращаем Истина иначе, Ложь
    If shpName1 = shpName2 Then
        SelfWaterGetCheck = True
    Else
        SelfWaterGetCheck = False
    End If
End Function


'----------------------------------------Служебные Функции-----------------------------------------------
Private Function f_IdentShape(ByVal ai_ShapeIP As Integer) As Integer
'Функция идентифициурет фигуру и возвращает значение её типа
Dim Arr_PTVs(26, 1) As Integer
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
    Arr_PTVs(5, 0) = 40  'Гидроэлеватор
        Arr_PTVs(5, 1) = vb_ShapeType_GE
    Arr_PTVs(6, 0) = 42  'Разветвление
        Arr_PTVs(6, 1) = vb_ShapeType_Razv
    Arr_PTVs(7, 0) = 45  'Пеноподъемник
        Arr_PTVs(7, 1) = vb_ShapeType_PTV
    Arr_PTVs(8, 0) = 72  'Колонка
        Arr_PTVs(8, 1) = vb_ShapeType_PTV
    Arr_PTVs(9, 0) = 88  'Всасывающая линия с сеткой
        Arr_PTVs(9, 1) = vb_ShapeType_VsasSet
    Arr_PTVs(10, 0) = 100 'Напорная линия
        Arr_PTVs(10, 1) = vb_ShapeType_Hose
    Arr_PTVs(11, 0) = 101 'Всасывающая линия
        Arr_PTVs(11, 1) = vb_ShapeType_Hose
    Arr_PTVs(12, 0) = 1 'Автоцистерна пожарная
        Arr_PTVs(12, 1) = vb_ShapeType_Tech
    Arr_PTVs(13, 0) = 2 'АНР
        Arr_PTVs(13, 1) = vb_ShapeType_Tech
    Arr_PTVs(14, 0) = 20 'Рукавный автомобиль
        Arr_PTVs(14, 1) = vb_ShapeType_Tech
    Arr_PTVs(15, 0) = 161 'АЦЛ
        Arr_PTVs(15, 1) = vb_ShapeType_Tech
    Arr_PTVs(16, 0) = 162 'АЦКП
        Arr_PTVs(16, 1) = vb_ShapeType_Tech
    Arr_PTVs(17, 0) = 163 'АПП
        Arr_PTVs(17, 1) = vb_ShapeType_Tech
    Arr_PTVs(18, 0) = 8 'ПНС
        Arr_PTVs(18, 1) = vb_ShapeType_Tech
    Arr_PTVs(19, 0) = 9 'АА
        Arr_PTVs(19, 1) = vb_ShapeType_Tech
    Arr_PTVs(20, 0) = 20 'АР
        Arr_PTVs(20, 1) = vb_ShapeType_Tech
    Arr_PTVs(21, 0) = 13 'АГВТ
        Arr_PTVs(21, 1) = vb_ShapeType_Tech
    Arr_PTVs(22, 0) = 28 'мотопомпа
        Arr_PTVs(22, 1) = vb_ShapeType_Tech
    Arr_PTVs(23, 0) = 190 'Емкость с водой
        Arr_PTVs(23, 1) = vb_ShapeType_WaterContainer
    Arr_PTVs(24, 0) = 191 'Пенная вставка
        Arr_PTVs(24, 1) = vb_ShapeType_Tech  'Как ни странно, но так лучше
    Arr_PTVs(25, 0) = 10 'автомобиль пенного тушения
        Arr_PTVs(25, 1) = vb_ShapeType_Tech
    Arr_PTVs(26, 0) = 41 'пеносмеситель
        Arr_PTVs(26, 1) = vb_ShapeType_Tech  'Как ни странно, но так лучше
        
        
        
'---Указываем значение по умолчанию
    f_IdentShape = vb_ShapeType_Other

'---Проверяем является ли фигура
        For i = 0 To 26
            If ai_ShapeIP = Arr_PTVs(i, 0) Then f_IdentShape = Arr_PTVs(i, 1): Exit Function
        Next i

End Function


