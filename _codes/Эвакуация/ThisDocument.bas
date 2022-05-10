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

'������ ��� ������� �����
Public WithEvents CalcComBut As Office.CommandBarButton
Attribute CalcComBut.VB_VarHelpID = -1
'������ ��� ������������ ����� �����
Public WithEvents RenumComBut As Office.CommandBarButton
Attribute RenumComBut.VB_VarHelpID = -1
'������ ��� ������ ���� ����� �����
Public WithEvents SelectComBut As Office.CommandBarButton
Attribute SelectComBut.VB_VarHelpID = -1

'������ �� ���������� ��� ������������ ����������
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



'---------���� ������ � ������� �� ������ ������������
Public Sub ActivateToolbarButtons()
    Set CalcComBut = Application.CommandBars("���������").Controls("����������")
    Set RenumComBut = Application.CommandBars("���������").Controls("��������������")
    Set SelectComBut = Application.CommandBars("���������").Controls("������� ���")
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


'--------���� ������ � ������������--------------------
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
            '��������� ������ �������������� ����� ����������� �������� (���� ����� ��� �� �������):
            If Not ShapeHaveCell(conLine, "User.IndexPers") Then
                conLine.AddNamedRow visSectionUser, "IndexPers", 0
                SetCellVal conLine, "User.IndexPers", indexPers.ipEvacEdge
                SetCellVal conLine, "User.IndexPers.Prompt", "����� ���� ���������"
                SetCellVal conLine, "LineColor", 3
                SetCellVal conLine, "EndArrow", 13
                SetCellVal conLine, "EndArrowSize", 1
                SetCellVal conLine, "ShapeRouteStyle", 16
'                SetCellFrml conLine, "Rounding", "1000mm"
                
                '������ ��� ����������� ����� ������
                conLine.AddNamedRow visSectionProp, "EdgeLen", 0
                SetCellVal conLine, "Prop.EdgeLen.Label", "�����"
                SetCellFrml conLine, "EventXFMod", Replace("CallThis('GetShapeLen','���������')", "'", Chr(34))
                GetShapeLen conLine

            End If
            '��� ���������� ������ ���� (��� �������, ��� ��� �������������� ������) ��������� ����� �������������� �����, ��� ����� ����
            If Not ShapeHaveCell(shpFrom, "Prop.WayClass", "������� �����") Then
                SetCellFrml shpFrom, "Prop.WayLen", "Sheet." & conLine.ID & "!Prop.EdgeLen"
            End If
            ' ���� ���������� ����� - ������ �������� ������, �� ��������, ��� ��������� ��������� ����� �������� �������������� ����� - ������... ��������
            If ShapeHaveCell(shpFrom, "Prop.WayClass", "������� �����") Then
                SetCellVal conLine, "EndArrow", 0
            End If
            
        End If
    End If
End Sub
