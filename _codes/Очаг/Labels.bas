Attribute VB_Name = "Labels"
Sub InsertLabelSquare(ShpObj As Visio.Shape)
'��������� ���������� ������� ������� ������ ��������� � ������� �������
'---��������� ����������
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX
'---���������� ������ ������� �������
    '---���������� ����� � � � ��� ������
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.item("����.vss").Masters.ItemU("������� �������"), pntX, pntY)

'---���������� ��������� � ��������� ������ ������� � �������
    '---���������� ��������� � ��������� ������� ������� � �������
    Set mstrConnection = ThisDocument.Masters("���������")
    
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
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(255,0,0))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---��������� ����� ������� � ������� �������
    CellFormula = "Sheet." & ShpObj.ID & "!User.FireSquare"
    shpLabel.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = CellFormula
    CellFormula = "Sheet." & ShpObj.ID & "!User.ExtSquare"
    shpLabel.CellsSRC(visSectionUser, 1, visUserValue).FormulaU = CellFormula
    
'---�������� ���� ������������ ��������
    shpLabel.Cells("Prop.Square.Invisible").FormulaU = 1
   
'---������ ����� �� �������
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
'---���������� ��������
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelSquare"
End Sub


Public Sub SeekFire(ShpObj As Visio.Shape)
'��������� ��������� �������� ��������������� ���� � ���������� ��� ������ ����������� ��������������� ����
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim col As Collection

    On Error GoTo EX
'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
        If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.Cells("User.IndexPers") = 64 And OtherShape.HitTest(x, y, 0.01) > 1 Then
                ShpObj.Cells("Prop.FireSpeed").FormulaU = _
                 "Sheet." & OtherShape.ID & "!Prop.FireSpeedLine"
            End If
        End If
    Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    SaveLog Err, "SeekFire", ShpObj.Name
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

'Public Sub HideMaster()
'Dim mstrConnection As Visio.Master
'
'    Set mstrConnection = Application.Documents("�������������.vss").Masters("���������")
'
'    mstrConnection.Hidden = True
'
'End Sub
