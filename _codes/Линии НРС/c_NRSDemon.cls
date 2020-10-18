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
'-------------------ƒемон-контроллер насосно-рукавной системы-----------------------
Private NRSModel_ As c_NRSModel




Public Function Activate(ByRef shp As Visio.Shape) As c_NRSDemon

Set Activate = Me
End Function





Public Function ClearNRSModel() As c_NRSDemon
    NRSModel_.Clear
Set ClearNRSModel = Me
End Function
Public Function BuildNRSModel(ByRef shp As Visio.Shape) As c_NRSDemon
'—троим модель насосно-рукавной системы от переданной фигуры
Dim modelBuilder As c_NRSModelBuilder

'    On Error GoTo EX
'    ClearNRSModel
    
    Set modelBuilder = New c_NRSModelBuilder
        Set NRSModel_ = modelBuilder.Build(shp)
    Set modelBuilder = Nothing
    
    
    
    
    
    

Set BuildNRSModel = Me
Exit Function
EX:
    Set BuildNRSModel = Me
End Function


Public Function CalculateNRSModel() As c_NRSDemon
'–ассчитываем модель насосно-рукавной системы: ќсновна€ процедура
    On Error GoTo EX
    
    '÷икл расчетов:
        '–асчет от питателей к потребител€м
            CalculateNRSModelForward
        '–асчет от потребителей к питател€м
            CalculateNRSModelBackward
    
    
    
    
    
Set CalculateNRSModel = Me
Exit Function
EX:
    Set CalculateNRSModel = Me
End Function


Public Function CalculateNRSModelForward() As c_NRSDemon
'ќдин цикл пр€мого расчета модели насосно-рукавной системы
Dim node As c_NRSNode

    On Error GoTo EX
    
    For Each node In NRSModel_.StarterNodes
        CalculateNRSNodeForward node
    Next node
    
    
    
Set CalculateNRSModelForward = Me
Exit Function
EX:
    Set CalculateNRSModelForward = Me
End Function

Public Function CalculateNRSNodeForward(ByRef node As c_NRSNode) As c_NRSDemon
'ѕр€мой расчет узла насосно-рукавной системы
Dim nextNode As c_NRSNode
    
    Debug.Print node.ID
    
    On Error GoTo EX
    
    For Each nextNode In node.OutNodes
        nextNode.hIn = node.hOut
        CalculateNRSNodeForward nextNode
    Next nextNode
    
    
    
Set CalculateNRSNodeForward = Me
Exit Function
EX:
    Set CalculateNRSNodeForward = Me
End Function






Public Function CalculateNRSModelBackward() As c_NRSDemon
'ќдин цикл обратного расчета модели насосно-рукавной системы
    On Error GoTo EX
    
    
    
    
    
    
    
Set CalculateNRSModelBackward = Me
Exit Function
EX:
    Set CalculateNRSModelBackward = Me
End Function
