VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_HoseConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------Класс отслеживания соединений рукавных линий--------------------------------------------------
Option Explicit

'---------------------Коллекции-------------------------------------------------------------------------------------

'Public ColP_WStvols As New Collection
'Public Colp_TempCol As Collection

'---------------------Свойства--------------------------------------------------------------------------------------
Private WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Private cpO_InShape, cpO_OutShape As Visio.Shape 'Фигуры в которую входит поток и выходит, соответственно

'---Постоянные индеков графиков
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

'Const vb_LineTypeWork = 0                   'Рабочая рукавная линия
'Const vb_LineTypeMag = 1                    'Магистральная рукавная линия
'Const vb_LineTypeVsas = 2                   'Всасывающая рукавная линия
'Const vb_LineTypeMag66 = 3                  'Напорная рукавная линия 66мм
'Const vb_LineTypeNapVsas = 4                'Напорно-всасывающая рукавная линия 77мм
'Const vb_LineTypeShortHoses = 5             'Напорная линия составленная из рукавов диаметром 77мм и длиной по 4м

'Const csi_PodIndex = 5
'Const csi_PodEffIndex = 6





'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()

    Set vO_App = Visio.Application

End Sub


Private Sub Class_Terminate()
    Set vO_App = Nothing
End Sub





'-------------------Процедуры отслеживающие соединения и выполняющие соответствующие действия------------------------------------------------------------
Private Sub vO_App_ConnectionsAdded(ByVal Connects As IVConnects)
'Процедура отслеживает появление соединений фигур и запускает процедуру
    
    '---Проверяем какие фигуры соединены
    TurnIntoHoses Connects
    
    '---Соединяем фигуры
    Ps_ConnectionAdd Connects(1)

End Sub



Public Sub Ps_ConnectionAdd(ByRef aO_Conn As Visio.Connect)
'Процедура непосредственно осуществляет соединение фигур
Dim vO_FromShape, vO_ToShape As Visio.Shape
Dim vi_InRowNumber, vi_OutRowNumber As Integer

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

    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
Exit Sub

EndSub:
    Resume Next
'    Debug.Print "Принимающая фигура: " & cpO_InShape.Name
'    Debug.Print "Подающая фигура: " & cpO_OutShape.Name
'    Debug.Print "Фигура рукава: " & cpO_HoseShape.Name
'    Debug.Print Left(Connects(1).FromCell.Name, 18) & " -> " & Left(Connects(1).ToCell.Name, 18)
'    Debug.Print vO_FromShape & " -> " & vO_ToShape
'    Debug.Print Err.Description
    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
    SaveLog Err, "Ps_ConnectionAdd"
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
'            cpO_OutShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1
'            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Дополнение для лафетных водяных стволов и пенного ствола:
            If vi_IPInShape = 36 Or vi_IPInShape = 37 Or vi_IPInShape = 39 Or vi_IPInShape = 35 Then
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
        If (vb_OutShapeType = vb_ShapeType_PTV Or vb_OutShapeType = vb_ShapeType_Razv Or vb_OutShapeType = vb_ShapeType_GE) And vb_InShapeType = vb_ShapeType_Hose Then
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
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.ArrivalTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Указываем, что к фигуре подсоединены рукава на вход
            If cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 1
            ElseIf cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 2
            End If
            '---Проверяем, не получает ли рукав воду от того же МСП, которому отдает
            If SelfWaterGetCheck(cpO_OutShape) Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 0
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
            If cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 1
            ElseIf cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 2
            End If
        End If
    '---Рукав->Емкость
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_WaterContainer Then
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 1
        End If
    '---Емкость->Рукав
        If vb_OutShapeType = vb_ShapeType_WaterContainer And vb_InShapeType = vb_ShapeType_Hose Then
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
    
    
    
    
Exit Sub
EX:
    Debug.Print Err.description
'    Debug.Print vs_Formula
    SaveLog Err, "ps_LinkShapes"

End Sub

Private Sub vO_App_ConnectionsDeleted(ByVal Connects As IVConnects)
'Событие происходит в момент разъединения рукавов
'MsgBox "Произошел разрыв соединения!"
    Ps_ConnectionDelete Connects(1)
End Sub

Public Sub Ps_ConnectionDelete(ByRef aO_Conn As Visio.Connect)
'Процедура непосредственно осуществляет очистку соединения фигур
Dim vO_FromShape, vO_ToShape As Visio.Shape
Dim vi_InRowNumber, vi_OutRowNumber As Integer

On Error GoTo EX
    
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
    
    
    

    '---Запускаем процедуру отвязывания данных в фигурах
       ps_DisLinkShapes vi_InRowNumber, vi_OutRowNumber

    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
Exit Sub

EX:
'    Resume Next
    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
    SaveLog Err, "Ps_ConnectionDelete"
End Sub

Private Sub ps_DisLinkShapes(ByVal ai_InRowNumber As Integer, ByVal ai_OutRowNumber As Integer)
'Внутренняя процедура - удаляет связанные данные в разъединяемых фигурах
Dim vi_IPInShape, vi_IPOutShape As Integer
Dim vb_InShapeType, vb_OutShapeType As Byte
Dim i As Integer

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
        
'---В зависимости от типа разъединяемых фигур выбрать процедуру разъединения
    '---Рукав->ПТВ
        If vb_OutShapeType = vb_ShapeType_Hose And (vb_InShapeType = vb_ShapeType_PTV Or vb_InShapeType = vb_ShapeType_Razv Or vb_InShapeType = vb_ShapeType_GE) Then
            cpO_OutShape.Cells("Scratch.A1").Formula = 0
            cpO_OutShape.Cells("Scratch.B1").Formula = 0
            cpO_OutShape.Cells("User.FlowToShape").Formula = ""
'            cpO_OutShape.Cells("Prop.LineTime").Formula = cpO_OutShape.Cells("Prop.LineTime")
            '---Дополнение для лафетных водяных стволов и пенного ствола:
            If vi_IPInShape = 36 Or vi_IPInShape = 37 Or vi_IPInShape = 39 Or vi_IPInShape = 35 Then
            '---Указываем, что от фигуры отсоединяются рукава на вход
            cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 0
            End If
            '---Привязываем маневренность линии к маневренности ствола
'            If cpO_InShape.CellExists("Actions.MainManeure", 0) = True Then
                cpO_OutShape.Cells("Prop.ManeverHose").FormulaU = "Index(0, Prop.ManeverHose.Format)"
'            End If
        End If
    '---ПТВ->Рукав
        If (vb_OutShapeType = vb_ShapeType_PTV Or vb_OutShapeType = vb_ShapeType_Razv Or vb_OutShapeType = vb_ShapeType_GE) And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = 0
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = 0
            cpO_InShape.Cells("User.FlowFromShape").Formula = ""
        End If
    '---Рукав->Рукав
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.Cells("Scratch.A1").Formula = 0
            cpO_OutShape.Cells("Scratch.B1").Formula = 0
            cpO_OutShape.Cells("User.FlowToShape").Formula = ""
            cpO_InShape.Cells("User.FlowFromShape").Formula = ""
        End If
    '---Рукав->Техника основная
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Tech Then
            cpO_OutShape.Cells("Scratch.A1").Formula = 0
            cpO_OutShape.Cells("Scratch.B1").Formula = 0
            cpO_OutShape.Cells("User.FlowToShape").Formula = ""
'            cpO_OutShape.Cells("Prop.LineTime").Formula = cpO_OutShape.Cells("Prop.LineTime")
'            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---Указываем, что к фигуре подсоединены рукава на вход
            cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 0
        End If
    '---Техника основная->Рукав
        If vb_OutShapeType = vb_ShapeType_Tech And vb_InShapeType = vb_ShapeType_Hose Then
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = 0
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = 0
            cpO_InShape.Cells("User.FlowFromShape").Formula = ""
            '---Указываем, что к фигуре подсоединены рукава на выход
            cpO_OutShape.Cells("Scratch.A" & CStr(ai_OutRowNumber + 1)).Formula = 0
        End If
    '---Всасывающая сетка->Техника основная
    If vb_OutShapeType = vb_ShapeType_VsasSet And vb_InShapeType = vb_ShapeType_Tech Then
        cpO_OutShape.Cells("Scratch.A1").Formula = 0
        cpO_OutShape.Cells("Scratch.B1").Formula = 0
        '---Указываем, что к фигуре подсоединены рукава на вход
        cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 0
    End If
    '---Рукав->Емкость
    If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_WaterContainer Then
        cpO_OutShape.Cells("Scratch.A1").Formula = 0
        cpO_OutShape.Cells("Scratch.B1").Formula = 0
        cpO_OutShape.Cells("User.FlowToShape").Formula = ""
        cpO_InShape.Cells("Scratch.A" & CStr(ai_InRowNumber + 1)).Formula = 0
    End If
    '---Емкость->Рукав
    If vb_OutShapeType = vb_ShapeType_WaterContainer And vb_InShapeType = vb_ShapeType_Hose Then
        cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 2).Formula = 0
        cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 3).Formula = 0
        cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = 0
        cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = 0
        cpO_InShape.Cells("User.FlowFromShape").Formula = ""
    End If
    
    
    
Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "ps_DisLinkShapes"

End Sub

Private Sub vO_App_BeforeShapeDelete(ByVal Shape As IVShape)
'Событие происходит при удалении фигуры - очищаются все существовавшие соединения для данной фигуры
Dim vO_Conn As Visio.Connect
Dim i As Integer

'MsgBox Shape.Connects.Count

    '---Очищаем соединения для указанной фигуры
    For Each vO_Conn In Shape.Connects
        Ps_ConnectionDelete vO_Conn
    Next vO_Conn

End Sub


Private Sub s_ReconnectAll(ByRef a_Shape As Visio.Shape)
'---Процедура осуществляет полное обновлениевсех связей для указанной фигуры
    
    
    
    
End Sub

'Public Sub S_ReconnectHose(ai_ShapeID As Integer)
''Процедура осуществляет реконнект соединений для данной фигуры рукавной линии
'Dim vO_Conn As Visio.Connect
'Dim vO_ShapeCur As Visio.Shape
'
'    '---Определяем фигуру рукавной линии
'    Set vO_ShapeCur = Application.ActivePage.Shapes.ItemFromID(vO_ShapeCur)
'
'    '---Очищаем соединения для указанной фигуры
'    For Each vO_Conn In ShpObj.Connects
''        vO_Conn.del
'    '    MsgBox Conn.FromSheet & " <- " & Conn.ToSheet
'    Next Conn
'
'
'End Sub


'----------------------------------------Служебные процедуры-----------------------------------------------
Private Function SelfWaterGetCheck(ByRef hoseShp As Visio.Shape) As Boolean
'Прока проверяет, не осуществляется ли самозабор и если осуществляется, подключение к всасывающим патрубкам МСП обнуляется
Dim con As Visio.Connect
Dim shpName1 As String
Dim shpName2 As String
    
'---Проверяем какие фигуры подключены к концам рукава
    For Each con In hoseShp.Connects
'        Debug.Print con.ToSheet.NameU & " - " & con.ToCell.Name
'        Debug.Print con.FromSheet.NameU & " - " & con.FromCell.Name
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
Dim Arr_PTVs(27, 1) As Integer
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
    Arr_PTVs(27, 0) = 22 'Гребенка пеногенераторов
        Arr_PTVs(27, 1) = vb_ShapeType_PTV
        
        
        
'---Указываем значение по умолчанию
    f_IdentShape = vb_ShapeType_Other

'---Проверяем является ли фигура
        For i = 0 To 27
            If ai_ShapeIP = Arr_PTVs(i, 0) Then f_IdentShape = Arr_PTVs(i, 1): Exit Function
        Next i

End Function



'----------------------------------------------Проки автоматического соединения фигур------------------------------------------------
Private Sub TurnIntoHoses(ByRef Connects As IVConnects)
'Процедура обращения в рукавную линию (при соединении)
Dim cnct As Visio.Connect
Dim hoseShp As Visio.Shape
    
    On Error GoTo EndSub
    '---Определяем какая именно фигура является фигурой линии
    Set hoseShp = GetHoseShape(Connects)
    '---Если фигура рукава отсутствует или уже обращена - выходим
    If hoseShp Is Nothing Then Exit Sub
        
    '---проверяем, имеет ли фигура ДВЕ точки соединения
    If hoseShp.Connects.Count <> 2 Then Exit Sub
        
    '---Проверяем Является ли фигура линией
    If hoseShp.AreaIU > 0 Then Exit Sub
        
    '---Проверяем, не является ли выбранная фигура уже рукавом или другой фигурой с назначенными свойствами
    If hoseShp.RowCount(visSectionUser) > 0 Then Exit Sub
    
    '---Основная процедура обращения
    TurnIntoHosesMain hoseShp
      
Exit Sub
EndSub:
    SaveLog Err, "TurnIntoHoses"
End Sub

Private Sub TurnIntoHoses2(ByRef droppedShape As Visio.Shape)
'Процедура обращения в рукавную линию (при рисовании линии)
'Dim cnct As Visio.Connect
Dim hoseShp As Visio.Shape
    
'    '---Определяем является ли вброшенная фигура фигурой линии
'    Set hoseShp = CheckHoseShape(droppedShape)
'    '---Если фигура рукава отсутствует или уже обращена - выходим
'    If hoseShp Is Nothing Then Exit Sub
        
    On Error GoTo EndSub
    
    Set hoseShp = droppedShape
        
    '---проверяем, имеет ли фигура ДВЕ точки соединения
    If hoseShp.Connects.Count <> 2 Then Exit Sub
        
    '---Проверяем Является ли фигура линией
    If hoseShp.AreaIU > 0 Then Exit Sub
        
    '---Проверяем, не является ли выбранная фигура уже рукавом или другой фигурой с назначенными свойствами
    If hoseShp.RowCount(visSectionUser) > 0 Then Exit Sub
    
    '---Основная процедура обращения
    TurnIntoHosesMain hoseShp
    
Exit Sub
EndSub:
    SaveLog Err, "TurnIntoHoses2"
End Sub

Private Sub TurnIntoHosesMain(hoseShp As Visio.Shape)
'Основная процедура обращения
Dim cnct As Visio.Connect
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
    
    On Error GoTo EndSub
    
    '---Получаем тип первой фигуры
    Set cnct = hoseShp.Connects(1)
    Set shp1 = cnct.ToSheet
    '---Получаем тип второй фигуры
    Set cnct = hoseShp.Connects(2)
    Set shp2 = cnct.ToSheet
    
    '---Проверяем обе ли соединенные фигуры - фигуры ГраФиС
    If Not IsGraFisShapes(shp1, shp2) Then Exit Sub

    '---Получаем тип линии
    GetLineType shp1, shp2, hoseShp

Exit Sub
EndSub:
    SaveLog Err, "TurnIntoHosesMain"
End Sub

Private Function GetHoseShape(ByRef Connects As IVConnects) As Visio.Shape
Dim cnct As Visio.Connect
Dim hoseShp As Visio.Shape
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim shpType1 As Integer
Dim shpType2 As Integer
    
    On Error GoTo EndSub
    '---Получаем первую фигуру
    Set cnct = Connects(1)
    Set shp1 = cnct.ToSheet
    '---Получаем вторую фигуру
    Set cnct = Connects(1)
    Set shp2 = cnct.FromSheet
    
    '---Получаем типы фигур
    shpType1 = GetShapeType(shp1)
    shpType2 = GetShapeType(shp2)
    If shpType1 = 0 Xor shpType2 = 0 Then
        If shpType1 = 0 Then Set hoseShp = shp1
        If shpType2 = 0 Then Set hoseShp = shp2
    Else
        Set GetHoseShape = Nothing
    End If
    Set GetHoseShape = hoseShp
    
Exit Function
EndSub:
'    Set hoseShp = Nothing
    SaveLog Err, "GetHoseShape"
End Function

Private Function GetShapeType(ByRef shp As Visio.Shape) As Integer
'Функция возвращает тип фигуры - если фигура не фигура ГраФиС - возвращается vb_ShapeType_Other
    If shp.CellExists("User.IndexPers", 0) Then
        GetShapeType = f_IdentShape(shp.Cells("User.IndexPers").Result(visNumber))
    Else
        GetShapeType = vb_ShapeType_Other
    End If
End Function

Private Function IsSpecificShape(ByRef shp As Visio.Shape, ByVal indexPers As Integer) As Boolean
'Функция возвращает истина, если указанная фигура имеет указанный IndexPers - если не ГраФиС, возвращается Ложь
    If shp.CellExists("User.IndexPers", 0) Then
        IsSpecificShape = (shp.Cells("User.IndexPers").Result(visNumber) = indexPers)
    Else
        IsSpecificShape = False
    End If
End Function


Private Sub GetLineType(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, _
                     ByRef hoseShape As Visio.Shape)
'Функция возвращает тип необходимой линии
Dim shp1Type As Integer
Dim shp2Type As Integer
Dim diameter As Integer
    
    On Error GoTo EndSub
    
    'Получаем тип первой фигуры
    shp1Type = GetShapeType(shp1)
    'Получаем тип второй фигуры
    shp2Type = GetShapeType(shp2)
    
    '---Проверяем кикие именно фигуры соединены и делаем вывод о типе линии
    '---Между ПТВ
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_PTV, vb_ShapeType_PTV) Then
        '---Проверяем какой условный проход имеет ПТВ (применяется для стволов)
        diameter = GetDiameter(shp1, shp2)
        If diameter = 0 Then
            MakeHoseLine 51, 0
        Else
            MakeHoseLine diameter, 0
        End If
        Exit Sub
    End If
    '---Между ПТВ и разветвлением
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_PTV, vb_ShapeType_Razv) Then
        diameter = GetDiameter(shp1, shp2)
        If diameter = 0 Then
            MakeHoseLine 51, 0
        Else
            MakeHoseLine diameter, 0
        End If
        Exit Sub
    End If
    '---Между ПТВ и Техникой
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_PTV, vb_ShapeType_Tech) Then
        diameter = GetDiameter(shp1, shp2)
        If diameter = 0 Then
            '---Проверяем не колонка ли подключена
            If shp1.Cells("User.IndexPers").Result(visNumber) = 72 Or _
                    shp2.Cells("User.IndexPers").Result(visNumber) = 72 Then
                '---Проверяем к какой точке колонки подключена линия
                If IsColonkaLineRight(shp1, shp2, hoseShape) Then
                    MakeNapVsasHoseLine
                Else
                    MakeHoseLine 77, 2
                End If
            Else
                MakeHoseLine 51, 0
            End If
        Else
            MakeHoseLine diameter, 0
        End If

        Exit Sub
    End If
    '---Между Техникой и разветвлениями
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Tech, vb_ShapeType_Razv) Then
        MakeHoseLine 77, 1
        Exit Sub
    End If
    '---Между разветвлениями и разветвлениями
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Razv, vb_ShapeType_Razv) Then
        MakeHoseLine 77, 1
        Exit Sub
    End If
    '---Между техникой и техникой
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Tech, vb_ShapeType_Tech) Then
        MakeHoseLine 77, 1
        Exit Sub
    End If
    '---Между техникой и емкостями
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Tech, vb_ShapeType_WaterContainer) Then
        '---Проверяем к какой точке емкости подключена линия
            If IsEmkostIn(shp1, shp2, hoseShape) Then
                MakeHoseLine 77, 1
            Else
                MakeVHoseLine
            End If
        Exit Sub
    End If
    '---Между техникой и гидроэлеваторами
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Tech, vb_ShapeType_GE) Then
        '---Проверяем к какой точке гидроэлеватора подключена линия
        If IsGELineIn(shp1, shp2, hoseShape) Then
            MakeHoseLine 66, 1
        Else
            MakeHoseLine 77, 1
        End If
        Exit Sub
    End If
    '---Между разветвлениями и гидроэлеваторами
    If GetOrEqual(shp1Type, shp2Type, vb_ShapeType_Razv, vb_ShapeType_GE) Then
        '---Проверяем к какой точке гидроэлеватора подключена линия
        If IsGELineIn(shp1, shp2, hoseShape) Then
            MakeHoseLine 66, 1
        Else
            MakeHoseLine 77, 1
        End If
        Exit Sub
    End If

Exit Sub
EndSub:
    SaveLog Err, "GetLineType"
End Sub

Private Function IsGraFisShapes(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape) As Boolean
'Функция возвращет Истина, если фигуры явля.тся фигурой ГраФиС
    IsGraFisShapes = shp1.CellExists("IndexPers", 0) = 0 And shp2.CellExists("IndexPers", 0) = 0
End Function

Private Function GetOrEqual(ByVal val1 As Integer, ByVal val2 As Integer, _
                            ByVal type1 As Integer, ByVal type2 As Integer) As Boolean
'Функция возвращает Истниу, если в первой паре указанных значений встречается комбинация второй пары
    If (val1 = type1 And val2 = type2) Or (val1 = type2 And val2 = type1) Then
        GetOrEqual = True
    Else
        GetOrEqual = False
    End If
End Function

Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
'Событие добавления новой фигуры
    '---Пытаемся обратить доавленную фигуру в рукавную линию
    TurnIntoHoses2 Shape
End Sub

Private Function IsGELineIn(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, _
                            ByRef hoseShape As Visio.Shape) As Boolean
'Функция возвращает Истина, если одна из фигур - Гидроэлеватор и к его точке входа подключена рукавная линия
Dim shpGE As Visio.Shape
Dim cnct As Visio.Connect

    '---Проверяем какая именно из фигур - гидроэлеватор
    If GetShapeType(shp1) = vb_ShapeType_GE Then Set shpGE = shp1
    If GetShapeType(shp2) = vb_ShapeType_GE Then Set shpGE = shp2
    
    If hoseShape.Connects(1).ToSheet = shpGE Then Set cnct = hoseShape.Connects(1)
    If hoseShape.Connects(2).ToSheet = shpGE Then Set cnct = hoseShape.Connects(2)
    
    IsGELineIn = Left(cnct.ToCell.Name, 18) = "Connections.GFS_In"
End Function

Private Function IsColonkaLineRight(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, _
                            ByRef hoseShape As Visio.Shape) As Boolean
'Функция возвращает Истина, если одна из фигур - Пожарная колонка и к его правому патрубку подключена рукавная линия
Dim shpColonka As Visio.Shape
Dim cnct As Visio.Connect

    '---Проверяем какая именно из фигур - колонка
    If IsSpecificShape(shp1, 72) Then Set shpColonka = shp1
    If IsSpecificShape(shp2, 72) Then Set shpColonka = shp2
    
    If hoseShape.Connects(1).ToSheet = shpColonka Then Set cnct = hoseShape.Connects(1)
    If hoseShape.Connects(2).ToSheet = shpColonka Then Set cnct = hoseShape.Connects(2)
    
    IsColonkaLineRight = Left(cnct.ToCell.Name, 20) = "Connections.GFS_Out2"
End Function

Private Function IsEmkostIn(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, _
                            ByRef hoseShape As Visio.Shape) As Boolean
'Функция возвращает Истина, если одна из фигур - Емкость и к ее входящейму коннекшн поинту подключена рукавная линия
Dim shpEmkost As Visio.Shape
Dim cnct As Visio.Connect

    '---Проверяем какая именно из фигур - колонка
    If IsSpecificShape(shp1, 190) Then Set shpEmkost = shp1
    If IsSpecificShape(shp2, 190) Then Set shpEmkost = shp2
    
    If hoseShape.Connects(1).ToSheet = shpEmkost Then Set cnct = hoseShape.Connects(1)
    If hoseShape.Connects(2).ToSheet = shpEmkost Then Set cnct = hoseShape.Connects(2)
    
    IsEmkostIn = Left(cnct.ToCell.Name, 18) = "Connections.GFS_In"
End Function

Private Function GetDiameter(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape) As Integer
'Функция возвращает диаметр рукава для подключения к стволу
Dim diameterUsl As Integer
    
    On Error GoTo EndSub
    
    If shp1.CellExists("User.DiameterIn", 0) Then
        diameterUsl = shp1.Cells("User.DiameterIn").Result(visNumber)
    End If
    If shp2.CellExists("User.DiameterIn", 0) Then
        diameterUsl = shp2.Cells("User.DiameterIn").Result(visNumber)
    End If
    
    '---Проверяем получено ли значение диаметра
    If diameterUsl = 0 Then
        GetDiameter = diameterUsl      'Если нет - выходим, возвращая 0
        Exit Function
    End If
    
    '---Преобразуем условные проходы в реальные диаметры рукавов
    Select Case diameterUsl
        Case Is = 25
            GetDiameter = 25
        Case Is = 40
            GetDiameter = 38
        Case Is = 50
            GetDiameter = 51
        Case Is = 60
            GetDiameter = 66
        Case Is = 70
            GetDiameter = 77
    End Select
    
Exit Function
EndSub:
    SaveLog Err, "GetDiameter"
End Function
