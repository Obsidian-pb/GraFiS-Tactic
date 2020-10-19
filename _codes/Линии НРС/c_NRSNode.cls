VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------Узел насосно-рукавной системы-----------------------
Private nodeShape_ As Visio.Shape                   'Фигура связанная с данным узлом
Private S_ As Double                                'Гидравлическое сопротивление (!!!для рукавов - ВСЕЙ ЛИНИИ)
Private Q_ As Single                                'Расход через узел
Private H_ As Single                                'Напор создаваемый самим элементом НРС (например, насосом АЦ)
Private P_ As Single                                'Коэффициент проводимости (для стволов в основном, но можно применять и для прочего оборудования)
Private Z_ As Single                                'Перепад высот (для линий)
Private hIn_ As Single                              'Напор на входах (максимален среди всех узлов на вход)
'Private hOut_ As Single                            'Напор на выходах (одинаков для всех выходов)
Private outNodes_ As Collection                     'Коллекция узлов подключенных на выход
Private inNodes_ As Collection                      'Коллекция узлов подключенных на вход
Private roundAccuracy_ As Byte                      'Точность округления
Private isStarter_ As Boolean                       'Питатель НРС
Private isEnder_ As Boolean                         'Потребитель НРС



Private Sub Class_Initialize()
    Set outNodes_ = New Collection
    Set inNodes_ = New Collection
End Sub
Private Sub Class_Terminate()
    Set outNodes_ = Nothing
    Set inNodes_ = Nothing
    Set nodeShape_ = Nothing
End Sub
Public Function Activate(ByRef shp As Visio.Shape, Optional ByVal a_S As Double = 0, Optional ByVal a_Q As Double = 0, Optional ByVal a_H As Single = 0, Optional ByVal a_P As Single = 1, _
                         Optional ByVal rAcc As Byte = 4) As c_NRSNode
    Set nodeShape_ = shp
    CheckStarterEnder
    
    S_ = a_S
    Q_ = a_Q
    H_ = a_H
    P_ = a_P
    
    roundAccuracy_ = rAcc
    
    FillProps
    
    
Set Activate = Me
End Function







Public Property Get S() As Double
    S = S_
End Property
Public Property Let S(ByVal vNewValue As Double)
    S_ = vNewValue
End Property

Public Property Get Q() As Single
    Q = Q_
End Property
Public Property Let Q(ByVal vNewValue As Single)
    Q_ = vNewValue
End Property
Public Function Q_Clear() As c_NRSNode
    Q_ = 0
Set Q_Clear = Me
End Function
Public Function Q_Raise(ByVal a_Q As Single) As c_NRSNode
    Q_ = Q_ + a_Q
Set Q_Raise = Me
End Function

Public Property Get hLost() As Double
    hLost = S_ * Q_ ^ 2
End Property

Public Property Get QIn() As Single
    QIn = Q_ / inNodes_.Count
End Property

Public Property Get QStv() As Single
'Производительности стволов (и возможно другого оборудования)
    QStv = P_ * Sqr(hOut)
End Property

'Public Property Get QOut() As Single
'Dim mx As Single
'Dim node As c_NRSNode
'
'    For Each node In OutNodes
'        if
'    Next node
'
'End Property




Public Property Get hIn() As Single
    hIn = hIn_
End Property
Public Property Let hIn(ByVal vNewValue As Single)
    If vNewValue > hIn_ Then hIn_ = vNewValue
End Property

Public Property Get hOut() As Single
'    hOut = hOut_
    hOut = H_ + hIn_ - hLost - Z_
End Property
'Public Property Let hOut(ByVal vNewValue As Single)
'    hOut_ = vNewValue
'End Property

Public Property Get OutNodes() As Collection
    Set OutNodes = outNodes_
End Property
'Public Property Let OutNodes(ByVal vNewValue As Single)
'    outNodes_ = vNewValue
'End Property
Public Function AddOutNode(ByRef node As c_NRSNode) As c_NRSNode
    AddUniqueCollectionItem outNodes_, node

Set AddOutNode = Me
End Function
Public Function DeleteOutNode(ByRef node As c_NRSNode) As c_NRSNode
    RemoveFromCollection outNodes_, node

Set DeleteOutNode = Me
End Function

Public Property Get InNodes() As Collection
    Set InNodes = inNodes_
End Property
Public Function AddInNode(ByRef node As c_NRSNode) As c_NRSNode
    AddUniqueCollectionItem inNodes_, node

Set AddInNode = Me
End Function
Public Function DeleteInNode(ByRef node As c_NRSNode) As c_NRSNode
    RemoveFromCollection inNodes_, node

Set DeleteInNode = Me
End Function

Public Property Get isStarter() As Boolean
    isStarter = isStarter_
End Property
Public Property Get isEnder() As Boolean
    isEnder = isEnder_
End Property

Public Property Get NodeShape() As Visio.Shape
    Set NodeShape = nodeShape_
End Property

Public Property Get ID() As String
    ID = CStr(nodeShape_.ID)
End Property


Private Sub CheckStarterEnder()
    isStarter_ = IsGFSShapeWithIP(nodeShape_, Array(ipAC, ipANR, ipPNS, ipAA, ipAV, ipAKT, ipAP, ipAGVT, ipPoezd, ipMotoPump, ipPricep, ipKorabl, ipKater, ipACL, ipACKP, ipAPP, ipKolonka))
    isEnder_ = IsGFSShapeWithIP(nodeShape_, Array(ipStvolRuch, ipStvolRuchPena, ipStvolLafVoda, ipStvolLafPena, ipStvolLafPoroshok, ipStvolLafVozimiy, ipPenopodemnik, ipStvolGas, ipStvolPoroshok))
End Sub


'---------------------------Большие инструментальные процедуры---------------------------
Public Function Realize() As c_NRSNode
'Здесь происходит указание текущих значений для фигур, по итогам расчета
    On Error GoTo EX
    Application.EventsEnabled = False
    
    'Для разных фигур разное
    Select Case CellVal(nodeShape_, "User.IndexPers")
        Case Is = indexPers.ipAC
            SetCellVal nodeShape_, "Prop.PodOut", Q_
'            SetCellVal nodeShape_, "Prop.PumpPressureCur", Q_
        Case Is = indexPers.ipRukavLineNapor
            SetCellVal nodeShape_, "Prop.Flow", Q_
            SetCellVal nodeShape_, "Prop.HeadInHose", hOut
            SetCellVal nodeShape_, "Prop.HeadLost", hLost
        Case Is = indexPers.ipRazvetvlenie
            SetCellVal nodeShape_, "Prop.StreamThrough", Q_
            SetCellVal nodeShape_, "Prop.HeadInRazv", hOut
        Case Is = indexPers.ipStvolRuch
            SetCellVal nodeShape_, "User.PodOut", QStv
            SetCellVal nodeShape_, "User.Head", hOut
            '!!!Не забыть учитывать подъем ствола: =GUARD(IF(Scratch.A1-Prop.StvolHeight<0,0,Scratch.A1-Prop.StvolHeight))
    End Select
    
    
    
Set Realize = Me
Exit Function
EX:
    Application.EventsEnabled = True
    Set Realize = Me
End Function

Private Sub FillProps()
'Заполняем свойства фигур в соответсвии с их IndexPers
Dim val1 As Variant
Dim val2 As Variant

    P_ = CellVal(nodeShape_, "User.ProvKoeff", , 1)

    If CellVal(nodeShape_, "Prop.ShowLenightDirect", visUnitsString) = "Да" Then
'            (Prop.HoseResistance                       / 20)                                    * Prop.LineLenightS                        '* Prop.Flow^2                            + Prop.Different
        S_ = (CellVal(nodeShape_, "Prop.HoseResistance") / CellVal(nodeShape_, "User.HoseLength")) * CellVal(nodeShape_, "Prop.LineLenightS") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    Else
'            Prop.HoseResistance                        * Prop.HoseNeed                        '* Prop.Flow^2                            + Prop.Different
        S_ = CellVal(nodeShape_, "Prop.HoseResistance") * CellVal(nodeShape_, "Prop.HoseNeed") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    End If
    
    If isEnder_ Then
        hIn = CellVal(nodeShape_, "Scratch.A1")
    ElseIf isStarter_ Then
        val1 = CellVal(nodeShape_, "Prop.PodOut")
        val2 = CellVal(nodeShape_, "Prop.Flow")
        If val1 > val2 Then Q_ = val1 Else Q_ = val2
        
        H_ = CellVal(nodeShape_, "Prop.PumpPressureCur")
    End If

    
    
    
    Z_ = CellVal(nodeShape_, "Prop.Different")
End Sub

