Attribute VB_Name = "Labels"
Sub InsertLabelName(ShpObj As Visio.Shape)
'��������� ���������� ������� �������� ��������� ������������� ��������� � ������� �������������
'---��������� ����������
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.cell, Cell2 As Visio.cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX

'---���������� ������ ������� �������
    '---���������� ����� � � � ��� ������
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������������.vss").Masters.ItemU("������� �������� �������������"), pntX, pntY)

'---���������� ��������� � ��������� ������ ������������� � �������
    '---���������� ��������� � ��������� ������� ������������� � �������
    Set mstrConnection = Application.Documents("�������������.vss").Masters("���������")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set vsoCell1 = shpConnection.CellsU("BeginX")
    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    Set vsoCell1 = shpConnection.CellsU("EndX")
    Set vsoCell2 = shpLabel.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2

'---���������� �������� ������ ����������
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(80,123,175))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---��������� ����� ������� � ������� �������������
    CellFormula = "Sheet." & ShpObj.ID & "!Prop.Name"
    shpLabel.CellsSRC(visSectionTextField, 0, visFieldCell).FormulaU = CellFormula

   
'---������ ����� �� �������
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelName"
End Sub

Sub InsertLabelValue(ShpObj As Visio.Shape)
'��������� ���������� ������� ������ ��������� ������������� ��������� � ������� �������������
'---��������� ����������
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.cell, Cell2 As Visio.cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX
'---���������� ������ ������ �������������
    '---���������� ����� � � � ��� ������
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.Item("�������������.vss").Masters.ItemU("����� ��������� �������������"), pntX, pntY)

'---���������� ��������� � ��������� ������ ������������� � �������
    '---���������� ��������� � ��������� ������� ������������� � �������
    Set mstrConnection = Application.Documents("�������������.vss").Masters("���������")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set vsoCell1 = shpConnection.CellsU("BeginX")
    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    Set vsoCell1 = shpConnection.CellsU("EndX")
    Set vsoCell2 = shpLabel.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2

'---���������� �������� ������ ����������
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(80,123,175))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---��������� ����� ������� � ������� �������������
    CellFormula = "Sheet." & ShpObj.ID & "!Prop.Value"
    shpLabel.CellsSRC(visSectionTextField, 0, visFieldCell).FormulaU = CellFormula

   
'---������ ����� �� �������
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelValue"
End Sub

Public Sub ConnectedShapesLostCheck(ShpObj As Visio.Shape)
'��������� ���������, �� ���� �� ������� ���� �� ����� ����������� �����������, � ���� ����, �� ������� ��� ���������
Dim CellsVal(4) As String
    
On Error GoTo EX
    
    CellsVal(0) = ShpObj.Cells("BegTrigger").FormulaU
    CellsVal(1) = ShpObj.Cells("BegTrigger").Result(visUnitsString)
    CellsVal(2) = ShpObj.Cells("EndTrigger").FormulaU
    CellsVal(3) = ShpObj.Cells("EndTrigger").Result(visUnitsString)
    
    If CellsVal(0) = CellsVal(1) Or CellsVal(2) = CellsVal(3) Then
        ShpObj.Delete
    End If
Exit Sub
EX:
    '������
End Sub
