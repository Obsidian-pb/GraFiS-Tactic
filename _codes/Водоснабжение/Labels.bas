Attribute VB_Name = "Labels"
Sub InsertLabelName(ShpObj As Visio.Shape)
'��������� ���������� ������� �������� ��������� ������������� ��������� � ������� �������������
'---��������� ����������
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX

'---���������� ������ ������� �������
    '---���������� ����� � � � ��� ������
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.drop(Application.Documents.Item("�������������.vss").Masters.ItemU("������� �������� �������������"), pntX, pntY)

'---���������� ��������� � ��������� ������ ������������� � �������
    '---���������� ��������� � ��������� ������� ������������� � �������
    Set mstrConnection = ThisDocument.Masters("���������")
    
    Set shpConnection = Application.ActiveWindow.Page.drop(mstrConnection, 2, 2)
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
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX
'---���������� ������ ������ �������������
    '---���������� ����� � � � ��� ������
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.drop(Application.Documents.Item("�������������.vss").Masters.ItemU("����� ��������� �������������"), pntX, pntY)

'---���������� ��������� � ��������� ������ ������������� � �������
    '---���������� ��������� � ��������� ������� ������������� � �������
    Set mstrConnection = ThisDocument.Masters("���������")
    
    Set shpConnection = Application.ActiveWindow.Page.drop(mstrConnection, 2, 2)
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
'        ShpObj.Cells("GlueType").FormulaU = 4
        FindAndDeleteLabels ShpObj
        ShpObj.Delete
    End If
Exit Sub
EX:
    '������
End Sub

Private Sub FindAndDeleteLabels(ByRef shp As Visio.Shape)
'���� � ������� ������� � ������� ��������� ������ ������ ����������
Dim con As Visio.Connect

    For Each con In shp.FromConnects
        If IsGFSShapeWithIP(con.ToSheet, 152) Then con.ToSheet.Delete
        If IsGFSShapeWithIP(con.FromSheet, 152) Then con.FromSheet.Delete
    Next con
    For Each con In shp.Connects
        If IsGFSShapeWithIP(con.ToSheet, 152) Then con.ToSheet.Delete
        If IsGFSShapeWithIP(con.FromSheet, 152) Then con.FromSheet.Delete
    Next con
End Sub



Function InsertDistance(ShpObj As Visio.Shape, Optional Contex As Integer = 0)
'��������� ���������� strelki rasstoiania � ������� �������������
'---��������� ����������
Dim shpTarget As Visio.Shape
Dim shpConnection As Visio.Shape, vsO_Shape As Visio.Shape
Dim mstrConnection As Visio.Master, mstrSrelka As Visio.Master
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
Dim vsi_ShapeIndex As Integer

vsi_ShapeIndex = 0

    On Error GoTo EX
    
    '---���������� ��� ������ � nahodim ochag
    For Each shpTarget In Application.ActivePage.Shapes
        If shpTarget.CellExists("User.IndexPers", 0) = True And shpTarget.CellExists("User.Version", 0) = True Then '�������� �� ������ ������� ������
            If shpTarget.Cells("User.Version") >= CP_GrafisVersion Then  '��������� ������ ������
                vsi_ShapeIndex = shpTarget.Cells("User.IndexPers")   '���������� ������ ������ ������
                If vsi_ShapeIndex = 64 Then
                    Exit For
                Else
                    If vsi_ShapeIndex = 70 Then
                    Exit For
                    Else
                        vsi_ShapeIndex = 0
                    End If
                End If
            End If
        End If
     Next
     
     If Contex = 0 And vsi_ShapeIndex = 0 Then Exit Function
    
'---���������� ��������� � ��������� ������ ������������� � ochag
    Set mstrConnection = ThisDocument.Masters("Distance")
    Set shpConnection = Application.ActiveWindow.Page.drop(mstrConnection, 2, 2)
  
    Set vsoCell1 = shpConnection.CellsU("BeginX")
    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    
    If vsi_ShapeIndex = 0 And Contex = 1 Then
        shpConnection.CellsU("EndX") = ShpObj.CellsU("pinX") + 550
        shpConnection.CellsU("EndY") = ShpObj.CellsU("pinY") + 550
    Else
        Set vsoCell1 = shpConnection.CellsU("EndX")
        Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    End If
    
''---�������� ������������� �������� ����� ��� ������������ ����� � ��������� �����������
    If ShpObj.CellExists("User.Distance", 0) = False Then ShpObj.AddNamedRow visSectionUser, "Distance", 0
    ShpObj.Cells("User.Distance").FormulaU = "Sheet." & shpConnection.ID & "!User.Length"
        
'---���������� �������� ������ ���������� i strelki
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(0,176,240))"
    shpConnection.Cells("Controls.ArrowPos").FormulaU = "IF(Width<6 m,Width*0.7*ThePage!User.GFS_Aspect,5 m*ThePage!User.GFS_Aspect)"
    shpConnection.Cells("Controls.TxtPos").FormulaU = "IF(STRSAME(Prop.ArrowStyle,""���������������� �����������""),IF(Width<6 m,Width*0.5*ThePage!User.GFS_Aspect,2.5 m*ThePage!User.GFS_Aspect),Width*0.5*ThePage!User.GFS_Aspect)"
    shpConnection.Cells("Controls.Ident").FormulaU = "IF(Width<6 m,Width*0.3*ThePage!User.GFS_Aspect,2 m *ThePage!User.GFS_Aspect)"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula
      
'---������ �����
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpObj, visSelect

Exit Function
EX:
    SaveLog Err, "InsertDistance"
End Function


