VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSModelControllerGidro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------Контроллер модели---------------------------------------------
'-------------------Реализован гидравлический механизм расчета--------------------
Private NRSModel_ As c_NRSModel

Private startersArray_ As Variant
Private enderArray_ As Variant

Private curNRSNode_ As c_NRSNode                        'Текущий узел с которым работает контроллер



Private Sub Class_Initialize()
    Set startersArray_ = Array(ipAC, ipANR, ipPNS, ipAA, ipAV, ipAKT, ipAP, ipAGVT, ipPoezd, ipMotoPump, ipPricep, ipKorabl, ipKater, ipACL, ipACKP, ipAPP, ipKolonka)
    Set enderArray_ = Array(ipStvolRuch, ipStvolRuchPena, ipStvolLafVoda, ipStvolLafPena, ipStvolLafPoroshok, ipStvolLafVozimiy, ipPenopodemnik, ipStvolGas, ipStvolPoroshok)
End Sub
Private Sub Class_Terminate()
    Set startersArray_ = Nothing
    Set enderArray_ = Nothing
    Set NRSModel_ = Nothing
End Sub
Public Function Activate(ByRef nrs_model As c_NRSModel) As C_NRSDebugger
    
    Set NRSModel_ = nrs_model
    
Set Activate = Me
End Function


Public Property Get NodeShape() As Visio.Shape
    Set NodeShape = curNRSNode_.NodeShape
End Property

Public Property Get ID() As String
    ID = CStr(curNRSNode_.NodeShape.ID)
End Property

'Вся механика изменения параметров модели НРС реализована в этом классе - в самой модели никакой логики быть не должно!
Public Function SetNRSNode(ByRef NRSNode As c_NRSNode) As c_NRSModelControllerGidro
    Set curNRSNode_ = NRSNode
Set SetNRSNode = Me
End Function
Public Function SetNRSNodeByID(ByVal NRSNodeID As String) As c_NRSModelControllerGidro
    Set curNRSNode_ = NRSModel_.GetNRSNodeByKey(NRSNodeID)
Set SetNRSNodeByID = Me
End Function




'----------------------Функции работы с текущим узлом---------------------
Public Function CalcH() As c_NRSModelControllerGidro
    curNRSNode_.H = Round(curNRSNode_.S * curNRSNode_.Q ^ 2, roundAccuracy_)
Set CalcH = Me
End Function

Public Function CalcQ() As c_NRSModelControllerGidro
    curNRSNode_.Q = curNRSNode_.Q / curNRSNode_.InNodes.Count
Set CalcQ = Me
End Function

Public Function CalcQStv() As c_NRSModelControllerGidro
'Производительности стволов (и возможно другого оборудования)
    curNRSNode_.Q = Round(curNRSNode_.P * Sqr(hOut), roundAccuracy_)
Set CalcQStv = Me
End Function

Public Function hLost() As Double
    hLost = curNRSNode_.S * curNRSNode_.Q ^ 2
End Function

Public Function hOut() As c_NRSModelControllerGidro
'    hOut = hOut_
'    Debug.Assert nodeShape_.ID <> 324
    curNRSNode_.hOut = curNRSNode_.H + curNRSNode_.hIn - hLost - Z_
Set c_NRSModelControllerGidro = Me
End Function

Public Function Q_Clear() As c_NRSModelControllerGidro
    curNRSNode_.Q = 0
Set Q_Clear = Me
End Function
Public Function Q_Raise(ByVal a_Q As Single) As c_NRSModelControllerGidro
    curNRSNode_.Q = curNRSNode_.Q + a_Q
Set Q_Raise = Me
End Function

Public Function CalchOut() As c_NRSModelControllerGidro
'    hOut = hOut_
'    ОСТАНОВИЛСЯ ЗДЕСЬ!
    hOut = curNRSNode_.H + curNRSNode_.hIn - curNRSNode_.hLost - curNRSNode_.Z
Set CalchOut = Me
End Function



Private Function CheckStarterEnder() As c_NRSModelControllerGidro
    curNRSNode_.isStarter = IsGFSShapeWithIP(curNRSNode_.NodeShape, startersArray_)
    curNRSNode_.isEnder = IsGFSShapeWithIP(curNRSNode_.NodeShape, enderArray_)

Set CheckStarterEnder = Me
End Function





'---------------------------Большие инструментальные процедуры---------------------------
Public Function Realize() As c_NRSModelControllerGidro
'Здесь происходит указание текущих значений для фигуры, по итогам расчета
    On Error GoTo ex
'    Application.EventsEnabled = False
    
    'Для разных фигур разное
    Select Case CellVal(curNRSNode_.NodeShape, "User.IndexPers")
        Case Is = indexPers.ipAC
            SetCellVal curNRSNode_.NodeShape, "Prop.PodOut", Q_
'            SetCellVal NRSNode.NodeShape, "Prop.PumpPressureCur", Q_
        Case Is = indexPers.ipRukavLineNapor
            SetCellVal curNRSNode_.NodeShape, "Prop.Flow", Q_
            SetCellVal curNRSNode_.NodeShape, "Prop.HeadInHose", hOut
            SetCellVal curNRSNode_.NodeShape, "Prop.HeadLost", hLost
        Case Is = indexPers.ipRazvetvlenie
            SetCellVal curNRSNode_.NodeShape, "Prop.StreamThrough", Q_
            SetCellVal curNRSNode_.NodeShape, "Prop.HeadInRazv", hOut
        Case Is = indexPers.ipStvolRuch
            SetCellVal curNRSNode_.NodeShape, "User.PodOut", QStv
            SetCellVal curNRSNode_.NodeShape, "User.Head", hOut
            '!!!Не забыть учитывать подъем ствола: =GUARD(IF(Scratch.A1-Prop.StvolHeight<0,0,Scratch.A1-Prop.StvolHeight))
    End Select
    
    
'Application.EventsEnabled = True
Set Realize = Me
Exit Function
ex:
'    Application.EventsEnabled = True
    Set Realize = Me
End Function

Private Function FillProps() As c_NRSModelControllerGidro
'Заполняем свойства фигур в соответсвии с их IndexPers
Dim val1 As Variant
Dim val2 As Variant

    P_ = CellVal(curNRSNode_.NodeShape, "User.ProvKoeff", , 1)

    If CellVal(curNRSNode_.NodeShape, "Prop.ShowLenightDirect", visUnitsString) = "Да" Then
'            (Prop.HoseResistance                                   / 20)                                                * Prop.LineLenightS                                   '* Prop.Flow^2                            + Prop.Different
        S_ = (CellVal(curNRSNode_.NodeShape, "Prop.HoseResistance") / CellVal(curNRSNode_.NodeShape, "User.HoseLength")) * CellVal(curNRSNode_.NodeShape, "Prop.LineLenightS") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    Else
'            Prop.HoseResistance                                   * Prop.HoseNeed                                   '* Prop.Flow^2                            + Prop.Different
        S_ = CellVal(curNRSNode_.NodeShape, "Prop.HoseResistance") * CellVal(curNRSNode_.NodeShape, "Prop.HoseNeed") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    End If
    
    If isEnder_ Then
        hIn = CellVal(curNRSNode_.NodeShape, "Scratch.A1")
    ElseIf isStarter_ Then
        val1 = CellVal(curNRSNode_.NodeShape, "Prop.PodOut")
        val2 = CellVal(curNRSNode_.NodeShape, "Prop.Flow")
        If val1 > val2 Then Q_ = val1 Else Q_ = val2
        
        H_ = CellVal(nodeShape_, "Prop.PumpPressureCur")
    End If

    
    
    
    Z_ = CellVal(curNRSNode_.NodeShape, "Prop.Different")
    
Set FillProps = Me
End Function



