VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------Демон-контроллер насосно-рукавной системы-----------------------
Private NRSModel_ As c_NRSModel

Public roundAccuracy As Byte              'Точность округления


Private Sub Class_Initialize()
    roundAccuracy = 2
End Sub
Private Sub Class_Terminate()
    Set NRSModel_ = Nothing
End Sub
'Public Function Activate(ByRef shp As Visio.Shape) As c_NRSDemon
'
'Set Activate = Me
'End Function








Public Function ClearNRSModel() As c_NRSDemon
    NRSModel_.Clear
Set ClearNRSModel = Me
End Function
Public Function BuildNRSModel(ByRef shp As Visio.Shape) As c_NRSDemon
'Строим модель насосно-рукавной системы от переданной фигуры
Dim modelBuilder As c_NRSModelBuilder

'    On Error GoTo EX
'    ClearNRSModel
    
    Set modelBuilder = New c_NRSModelBuilder
        Set NRSModel_ = modelBuilder.Build(shp)
    Set modelBuilder = Nothing
    
    
    
    
    
    

Set BuildNRSModel = Me
Exit Function
ex:
    Set BuildNRSModel = Me
End Function


Public Function CalculateNRSModel() As c_NRSDemon
'Рассчитываем модель насосно-рукавной системы: Основная процедура
Dim i As Integer
Const rounds = 50
    
    On Error GoTo ex
    
    Dim debugger As C_NRSDebugger
    Set debugger = New C_NRSDebugger
    debugger.Activate NRSModel_

    'Цикл расчетов:
        For i = 1 To rounds
            Debug.Print i
        'Обнуляем текущие напоры в узлах
            ClearhIn
'            debugger.PrintModelState
        'Расчет от питателей к потребителям
            CalculateNRSModelForward
'            debugger.PrintModelState
'!Обязательно обработать корректно!!!
'Это проблема невязки прямого и обратного расчета - при прямом изменяется напор но еще не изменился расход.
'При обратном уже изменился расход, но не изменился напор. Поэтому небольшая разница будет всегда! Обрабатывать ее нужно при помощи окурглений на мелких долях!!! 0.000001 >> 0.01
'            If i >= rounds Then Exit For
        'Обнуляем текущие расходы в узлах
            ClearQ
'            debugger.PrintModelState
        'Расчет от потребителей к питателям
            CalculateNRSModelBackward
'            debugger.PrintModelState
        Next i
        
        'Реализуем результаты расчета в фигурах
        Realize
    debugger.PrintModelState
    debugger.PrintQSum
    Set debugger = Nothing
    
Set CalculateNRSModel = Me
Exit Function
ex:
    Set CalculateNRSModel = Me
End Function


Public Function CalculateNRSModelForward() As c_NRSDemon
'Один цикл прямого расчета модели насосно-рукавной системы
Dim node As c_NRSNode

    On Error GoTo ex
    
    For Each node In NRSModel_.StarterNodes
        CalculateNRSNodeForward node
    Next node
    
    
    
Set CalculateNRSModelForward = Me
Exit Function
ex:
    Set CalculateNRSModelForward = Me
End Function

Public Function CalculateNRSNodeForward(ByRef node As c_NRSNode) As c_NRSDemon
'Прямой расчет узла насосно-рукавной системы
Dim nextNode As c_NRSNode
Dim tmpVal As Single
    
'    Debug.Print node.ID
    
    On Error GoTo ex
    
    For Each nextNode In node.OutNodes
'        Debug.Assert nextNode.ID <> 19
'        Debug.Assert nextNode.ID <> 67
'        Debug.Assert node.ID <> 19
'        Debug.Assert node.ID <> 324
        tmpVal = GetHOut(node)
        If tmpVal > nextNode.hIn Then nextNode.hIn = tmpVal     'Учитываем, что напор на входе равен максимальному напору из входящих
'        nextNode.hIn = GetHOut(node)    'node.hOut
'        Debug.Assert node.ID <> 324
        CalculateNRSNodeForward nextNode
    Next nextNode
    
    
    
Set CalculateNRSNodeForward = Me
Exit Function
ex:
    Set CalculateNRSNodeForward = Me
End Function


Public Function ClearQ() As c_NRSDemon
'Очищаем расходы во всех узлах кроме Enders
Dim node As c_NRSNode

    On Error GoTo ex

    For Each node In NRSModel_.NRSNodes
        If Not node.isEnder Then node.Q_Clear
    Next node
    
Set ClearQ = Me
Exit Function
ex:
    Set ClearQ = Me
End Function

Public Function ClearhIn() As c_NRSDemon
'Очищаем входящие напоры во всех узлах кроме starters (Возможнос)
Dim node As c_NRSNode

    On Error GoTo ex

    For Each node In NRSModel_.NRSNodes
'        If Not node.isStarter Then node.hIn_Clear
        node.hIn_Clear
    Next node
    
Set ClearhIn = Me
Exit Function
ex:
    Set ClearhIn = Me
End Function



Public Function CalculateNRSModelBackward() As c_NRSDemon
'Один цикл обратного расчета модели насосно-рукавной системы
Dim node As c_NRSNode

    On Error GoTo ex
    
    For Each node In NRSModel_.EnderNodes
        CalculateNRSNodeBackward node, GetQStv(node) 'node.QStv
    Next node
    
Set CalculateNRSModelBackward = Me
Exit Function
ex:
    Set CalculateNRSModelBackward = Me
End Function

Public Function CalculateNRSNodeBackward(ByRef node As c_NRSNode, ByRef a_Q As Double) As c_NRSDemon
'Обратный расчет узла насосно-рукавной системы
Dim prevNode As c_NRSNode
    
'    Debug.Print node.ID
    
'    On Error GoTo EX
    
    For Each prevNode In node.InNodes
'        If node.isEnder Then
'            prevNode.Q_Raise node.QStv
'            Debug.Print node.ID & " Q=" & node.QStv & ">>" & prevNode.ID & ": Q=" & prevNode.Q & "+" & node.QStv & "=" & prevNode.Q
'        Else
'            prevNode.Q_Raise node.Q
'            Debug.Print node.ID & " Q=" & node.Q & ">>" & prevNode.ID & ": Q=" & prevNode.Q & "+" & node.Q & "=" & prevNode.Q
'        End If
        prevNode.Q_Raise Round(a_Q / node.InNodes.Count, roundAccuracy)
        CalculateNRSNodeBackward prevNode, a_Q
    Next prevNode
    
    
    
Set CalculateNRSNodeBackward = Me
Exit Function
ex:
    Set CalculateNRSNodeBackward = Me
End Function


Public Function Realize() As c_NRSDemon
'Реализуем результаты расчета в фигурах
Dim node As c_NRSNode
    
'    Debug.Print node.ID
    
'    On Error GoTo EX
    
    For Each node In NRSModel_.NRSNodes
        RealizeNode node   'node.Realize
    Next node
    
Set Realize = Me
Exit Function
ex:
    Set Realize = Me
End Function



'-------------------------Блок работы с моделью и узлами------------------------------------------
'----------------------Функции работы с текущим узлом---------------------
Private Function GetHOut(ByRef node As c_NRSNode) As Single
    node.hOut = node.hIn + node.H - GetHLost(node)
    GetHOut = node.hOut
End Function

Private Function GetHLost(ByRef node As c_NRSNode) As Single
'    GetHLost = Round(node.S * node.Q ^ 2, roundAccuracy)
    node.hLost = Round(node.S * node.Q ^ 2 - node.Z, roundAccuracy)
    GetHLost = node.hLost
End Function

Private Function GetQStv(ByRef node As c_NRSNode) As Single
'Производительности стволов (и возможно другого оборудования)
    node.Q = Round(node.P * Sqr(node.hIn - GetHLost(node)), roundAccuracy)
    GetQStv = node.Q
End Function



'Public Function CalcH() As c_NRSDemon
'    curNRSNode_.H = Round(curNRSNode_.S * curNRSNode_.Q ^ 2, roundAccuracy_)
'Set CalcH = Me
'End Function
'
'Public Function CalcQ() As c_NRSModelControllerGidro
'    curNRSNode_.Q = curNRSNode_.Q / curNRSNode_.InNodes.Count
'Set CalcQ = Me
'End Function



'Public Function hLost() As Double
'    hLost = curNRSNode_.S * curNRSNode_.Q ^ 2
'End Function
'
'Public Function hOut() As c_NRSModelControllerGidro
''    hOut = hOut_
''    Debug.Assert nodeShape_.ID <> 324
'    curNRSNode_.hOut = curNRSNode_.H + curNRSNode_.hIn - hLost - Z_
'Set c_NRSModelControllerGidro = Me
'End Function
'
'Public Function Q_Clear() As c_NRSModelControllerGidro
'    curNRSNode_.Q = 0
'Set Q_Clear = Me
'End Function
'Public Function Q_Raise(ByVal a_Q As Single) As c_NRSModelControllerGidro
'    curNRSNode_.Q = curNRSNode_.Q + a_Q
'Set Q_Raise = Me
'End Function
'
'Public Function CalchOut() As c_NRSModelControllerGidro
''    hOut = hOut_
''    ОСТАНОВИЛСЯ ЗДЕСЬ!
'    hOut = curNRSNode_.H + curNRSNode_.hIn - curNRSNode_.hLost - curNRSNode_.Z
'Set CalchOut = Me
'End Function
'
'
'
'Private Function CheckStarterEnder() As c_NRSModelControllerGidro
'    curNRSNode_.isStarter = IsGFSShapeWithIP(curNRSNode_.NodeShape, startersArray_)
'    curNRSNode_.isEnder = IsGFSShapeWithIP(curNRSNode_.NodeShape, enderArray_)
'
'Set CheckStarterEnder = Me
'End Function





'---------------------------Большие инструментальные процедуры---------------------------
Public Function RealizeNode(ByRef node As c_NRSNode) As c_NRSDemon
'Здесь происходит указание текущих значений для фигуры, по итогам расчета
    On Error GoTo ex
'    Application.EventsEnabled = False
    
    'Для разных фигур разное
    Select Case CellVal(node.NodeShape, "User.IndexPers")
        Case Is = indexPers.ipAC
            SetCellVal node.NodeShape, "Prop.PodOut", node.Q
'            SetCellVal NRSNode.NodeShape, "Prop.PumpPressureCur", Q_
        Case Is = indexPers.ipRukavLineNapor
            SetCellVal node.NodeShape, "Prop.Flow", node.Q
            SetCellVal node.NodeShape, "Prop.HeadInHose", node.hOut
            SetCellVal node.NodeShape, "Prop.HeadLost", GetHLost(node)
        Case Is = indexPers.ipRazvetvlenie
            SetCellVal node.NodeShape, "Prop.StreamThrough", node.Q
            SetCellVal node.NodeShape, "Prop.HeadInRazv", GetHOut(node)
        Case Is = indexPers.ipStvolRuch
            SetCellVal node.NodeShape, "User.PodOut", GetQStv(node)
            SetCellVal node.NodeShape, "User.Head", GetHOut(node)
            '!!!Не забыть учитывать подъем ствола: =GUARD(IF(Scratch.A1-Prop.StvolHeight<0,0,Scratch.A1-Prop.StvolHeight))
    End Select
    
    
'Application.EventsEnabled = True
Set RealizeNode = Me
Exit Function
ex:
'    Application.EventsEnabled = True
    Set RealizeNode = Me
End Function







