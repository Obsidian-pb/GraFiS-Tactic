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
'-------------------���� �������-�������� �������-----------------------
Private nodeShape_ As Visio.Shape                   '������ ��������� � ������ �����
Private S_ As Double                                '�������������� ������������� (!!!��� ������� - ���� �����)
Private Q_ As Single                                '������ ����� ����
Private H_ As Single                                '����� ����������� ����� ��������� ��� (��������, ������� ��)
Private P_ As Single                                '����������� ������������ (��� ������� � ��������, �� ����� ��������� � ��� ������� ������������)
Private Z_ As Single                                '������� ����� (��� �����)
Private hIn_ As Single                              '����� �� ������ (���������� ����� ���� ����� �� ����)
Private hOut_ As Single                             '����� �� ������� (�������� ��� ���� �������)
Private hlost_ As Single                            '������ ������
Private outNodes_ As Collection                     '��������� ����� ������������ �� �����
Private inNodes_ As Collection                      '��������� ����� ������������ �� ����
Private isStarter_ As Boolean                       '�������� ���
Private isEnder_ As Boolean                         '����������� ���
Private prevQH(1) As Single                         '���������� �������� Q � hOut (���� �� ���� �� ����� ��� �����, �������� ���������� �������� Q)

'----------------������� ����-------------------------
'Public Event SChanged(ByVal a_S As Single)
'Public Event QChanged(ByVal a_Q As Single)
'Public Event HChanged(ByVal a_H As Single)
'Public Event PChanged(ByVal a_P As Single)
'Public Event ZChanged(ByVal a_Z As Single)
'Public Event hOutChanged(ByVal a_hOut As Single)
'Public Event hInChanged(ByVal a_hIn As Single)
'Public Event hLostChanged(ByVal a_hIn As Single)
Public Event PropChanged(ByVal val As Single, ByVal changedProp As NRSProp)


Private Sub Class_Initialize()
    Set outNodes_ = New Collection
    Set inNodes_ = New Collection
End Sub
Private Sub Class_Terminate()
    Set outNodes_ = Nothing
    Set inNodes_ = Nothing
    Set nodeShape_ = Nothing
End Sub
Public Function Activate(ByRef shp As Visio.Shape, Optional ByVal a_S As Double = 0, Optional ByVal a_Q As Double = 0, _
                         Optional ByVal a_H As Single = 0, Optional ByVal a_P As Single = 1) As c_NRSNode
    Set nodeShape_ = shp
'    CheckStarterEnder
    
    S_ = a_S
    Q_ = a_Q
    H_ = a_H
    P_ = a_P
    
'    roundAccuracy_ = rAcc
    
'    FillProps
    
    
Set Activate = Me
End Function







Public Property Get S() As Double
    S = S_
End Property
Public Property Let S(ByVal vNewValue As Double)
    S_ = vNewValue
    RaiseEvent PropChanged(S_, nrsPropS)
End Property

Public Property Get Q() As Single
    Q = Q_
End Property
Public Property Let Q(ByVal vNewValue As Single)
    Q_ = vNewValue
    RaiseEvent PropChanged(Q_, nrsPropQ)
End Property
Public Function Q_Clear() As c_NRSNode
    prevQH(0) = Q_
    Q_ = 0
Set Q_Clear = Me
End Function
Public Function Q_Raise(ByVal a_Q As Single) As c_NRSNode
    Q_ = Q_ + a_Q
Set Q_Raise = Me
End Function

Public Property Get H() As Single
    H = H_
End Property
Public Property Let H(ByVal vNewValue As Single)
    H_ = vNewValue
    RaiseEvent PropChanged(H_, nrsPropH)
End Property

Public Property Get P() As Single
    P = P_
End Property
Public Property Let P(ByVal vNewValue As Single)
    P_ = vNewValue
    RaiseEvent PropChanged(P_, nrsPropP)
End Property

Public Property Get Z() As Single
    Z = Z_
End Property
Public Property Let Z(ByVal vNewValue As Single)
    Z_ = vNewValue
    RaiseEvent PropChanged(Z_, nrsPropZ)
End Property

Public Property Get hOut() As Single
    hOut = hOut_
End Property
Public Property Let hOut(ByVal vNewValue As Single)
    hOut_ = vNewValue
    RaiseEvent PropChanged(hOut_, nrsProphOut)
End Property

Public Property Get hIn() As Single
    hIn = hIn_
End Property
Public Property Let hIn(ByVal vNewValue As Single)
    hIn_ = vNewValue
    RaiseEvent PropChanged(hIn_, nrsProphIn)
End Property
Public Function hIn_Clear() As c_NRSNode
    hIn_ = 0
Set hIn_Clear = Me
End Function

Public Property Get hLost() As Single
    hLost = hlost_
End Property
Public Property Let hLost(ByVal vNewValue As Single)
    hlost_ = vNewValue
    RaiseEvent PropChanged(hlost_, nrsProphLost)
End Property
Public Function hLost_Clear() As c_NRSNode
'    prevQH(1) = hlost_
    hlost_ = 0
Set hLost_Clear = Me
End Function

'Public Property Get hIn() As Single
'    hIn = hIn_
'End Property
'Public Property Let hIn(ByVal vNewValue As Single)
''    Debug.Assert nodeShape_.ID <> 213
'    If vNewValue > hIn_ Then hIn_ = vNewValue
'End Property
'Public Function hIn_Clear() As c_NRSNode
'    hIn_ = 0
'Set hIn_Clear = Me
'End Function


'Public Property Get hOut() As Single
''    hOut = hOut_
''    Debug.Assert nodeShape_.ID <> 324
'    hOut = H_ + hIn_ - hLost - Z_
'End Property


Public Property Get OutNodes() As Collection
    Set OutNodes = outNodes_
End Property
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
Public Property Let isStarter(ByVal vNewValue As Boolean)
    isStarter_ = vNewValue
End Property
Public Property Get isEnder() As Boolean
    isEnder = isEnder_
End Property
Public Property Let isEnder(ByVal vNewValue As Boolean)
    isEnder_ = vNewValue
End Property


Public Property Get NodeShape() As Visio.Shape
    Set NodeShape = nodeShape_
End Property

Public Property Get ID() As String
    ID = CStr(nodeShape_.ID)
End Property


Public Property Get PrevQ() As Single
    PrevQ = prevQH(0)
End Property











