VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Кнопка для расчета графа
Public WithEvents CalcComBut As Office.CommandBarButton
Attribute CalcComBut.VB_VarHelpID = -1
'Кнопка для переномрации узлов графа
Public WithEvents RenumComBut As Office.CommandBarButton
Attribute RenumComBut.VB_VarHelpID = -1
'Кнопка для выбора всех узлов графа
Public WithEvents SelectComBut As Office.CommandBarButton
Attribute SelectComBut.VB_VarHelpID = -1

'Ссылка на приложение для отслеживания соединений
Public WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1







Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    DeActivateApp
    DeActivateToolbarButtons
    RemoveTB_Evacuation
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    AddTB_Evacuation
    ActivateToolbarButtons
    ActivateApp
End Sub



'---------Блок работы с кнопкой на панели инстурментов
Public Sub ActivateToolbarButtons()
    Set CalcComBut = Application.CommandBars("Эвакуация").Controls("Рассчитать")
    Set RenumComBut = Application.CommandBars("Эвакуация").Controls("Перенумеровать")
    Set SelectComBut = Application.CommandBars("Эвакуация").Controls("Выбрать все")
End Sub
Public Sub DeActivateToolbarButtons()
    Set CalcComBut = Nothing
    Set RenumComBut = Nothing
    Set SelectComBut = Nothing
End Sub
Private Sub CalcComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    CalcTimes
End Sub
Private Sub RenumComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    RenumNodes
End Sub
Private Sub SelectComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    SelectNodes
End Sub


'--------Блок работы с соединениями--------------------
Public Sub ActivateApp()
    Set app = Visio.Application
End Sub
Public Sub DeActivateApp()
    Set app = Nothing
End Sub
Private Sub app_ConnectionsAdded(ByVal Connects As IVConnects)
Dim conLine As Visio.Shape
Dim shpFrom As Visio.Shape
Dim shpTo As Visio.Shape
    
    Set conLine = Connects.FromSheet
    If conLine.Connects.count = 2 Then
        Set shpFrom = conLine.Connects(1).ToSheet
        Set shpTo = conLine.Connects(2).ToSheet
        
        If IsGFSShapeWithIP(shpFrom, indexPers.ipEvacNode) And IsGFSShapeWithIP(shpTo, indexPers.ipEvacNode) Then
            Debug.Print shpFrom.Name & " --- " & shpTo.Name
            'Добавляем фигуре соединительной линии необходимые свойства (Если этого еще не сделано):
            If Not ShapeHaveCell(conLine, "User.IndexPers") Then
                conLine.AddNamedRow visSectionUser, "IndexPers", 0
                SetCellVal conLine, "User.IndexPers", indexPers.ipEvacEdge
                SetCellVal conLine, "User.IndexPers.Prompt", "Ребро пути эвакуации"
                SetCellVal conLine, "LineColor", 3
                SetCellVal conLine, "EndArrow", 13
                SetCellVal conLine, "EndArrowSize", 1
                SetCellVal conLine, "ShapeRouteStyle", 16
'                SetCellFrml conLine, "Rounding", "1000mm"
                
                'Ячейки для перерасчета длины фигуры
                conLine.AddNamedRow visSectionProp, "EdgeLen", 0
                SetCellVal conLine, "Prop.EdgeLen.Label", "Длина"
                SetCellFrml conLine, "EventXFMod", Replace("CallThis('GetShapeLen','Эвакуация')", "'", Chr(34))
                GetShapeLen conLine

            End If
            'Для предыдущей фигуры узла (при условии, что это горизонтальный проход) добавляем длину соединительной линии, как длину пути
            If Not ShapeHaveCell(shpFrom, "Prop.WayClass", "Дверной проем") Then
                SetCellFrml shpFrom, "Prop.WayLen", "Sheet." & conLine.ID & "!Prop.EdgeLen"
            End If
            ' Если предыдущая фигра - фигура дверного проема, то наоборот, для следующей указываем длину входящей соединительной линии - спорно... подумать
            If ShapeHaveCell(shpFrom, "Prop.WayClass", "Дверной проем") Then
                SetCellVal conLine, "EndArrow", 0
            End If
            
        End If
    End If
End Sub
