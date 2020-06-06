Attribute VB_Name = "m_WorkWithUniversalLabels"
Option Explicit


'------------------------������ ��� �������� �������� ������ � �������������� ���������----------


Public Sub SeekAnyGFSFigure(ShpObj As Visio.Shape)
'��������� ����� ��������� ��� ������ �� ����� � ���� ��� ������, �� ������� ��������
'������������� �������, ��������� ��
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim str As String
Dim delFlag As Boolean

    On Error GoTo EX
'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---���������� ��� ������ �� ��������
    delFlag = True
    For Each OtherShape In Application.ActivePage.Shapes
        If IsCorrectShapeForLabel(OtherShape) Then 'And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.HitTest(x, y, 10) >= 1 Then
            '---������������� ��������� ������ Prop.Property (�������� �������)
                ShpObj.Cells("Prop.Property.Format").Formula = """" & GetPropsList(OtherShape) & """" '"����������������"
            '---������������� ������ �� ������ Prop.Property ������
                str = GetPropsLinks(OtherShape)
                ShpObj.Cells("Prop.PropertyValue.Format").FormulaU = str
                
            '---���������� �������������� �����
                InsertLink OtherShape, ShpObj
                
            '---��������� ��� ������� ������ ����� ��������� �� �����
                delFlag = False
                
            '---������� �� �����
                Exit For
            End If
        End If
    Next OtherShape

'---� ������, ���� ������ �� ���� �� � ���� ��������� - ������� ��
    If delFlag Then
        ShpObj.Delete
        Exit Sub
    End If
    
'---���������� ��������
    On Error Resume Next
    Application.DoCmd (1312)

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    MsgBox "� ���� ���������� ��������� ������, ���� ��� ����� ����������� - ��������� � �������������!"
    SaveLog Err, "SeekAnyGFSFigure", ShpObj.Name
End Sub

Public Sub ConnectToGFSFigure(ShpObj As Visio.Shape)
'��������� ����� ��������� � ����� ������ ���� ���������
'������������� �������, ��������� ��
Dim OtherShape As Visio.Shape
Dim str As String

    On Error GoTo EX
    
'---���������� ������� ���������� � ������ ������������� �������, � ���� ����� 1 ��������� ������
    If ShpObj.Connects.Count = 1 Then
        Set OtherShape = ShpObj.Connects.Item(1).ToSheet
        GetCorrectShape OtherShape
    
    '---������������� ��������� ������ Prop.Property (�������� �������)
        ShpObj.Cells("Prop.Property.Format").Formula = """" & GetPropsList(OtherShape) & """" '"����������������"
    '---������������� ������ �� ������ Prop.Property ������
        str = GetPropsLinks(OtherShape)
        ShpObj.Cells("Prop.PropertyValue.Format").FormulaU = str
        
    Else
    '---� ������, ���� ���-�� ���������� �� ����� 1 ������ �������� �������� �� ���������
        ShpObj.Cells("Prop.Property.Format").Formula = """" & "����������������;������" & """"
        ShpObj.Cells("Prop.PropertyValue.Format").Formula = "Prop.Property.Prompt&" & Chr(34) & ";" & Chr(34) & _
            "&Prop.UserLabel.Prompt"
    End If

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    MsgBox "� ���� ���������� ��������� ������, ���� ��� ����� ����������� - ��������� � �������������!"
    SaveLog Err, "ConnectToGFSFigure", ShpObj.Name
End Sub

Private Function IsCorrectShapeForLabel(ShpTO As Visio.Shape)
'������� ���������, ����� �� ��������� ������� � ������ ������
IsCorrectShapeForLabel = True
'---�������� �� ������ �������
    If ShpTO.CellExists("User.IndexPers", 0) = True Then
        If ShpTO.Cells("User.IndexPers").Result(visNumber) = 152 Then
            IsCorrectShapeForLabel = False
        End If
    End If
'---�������� �� ������ ����� (1D)
    If InStr(1, ShpTO.Cells("Width").FormulaU, "SQRT") > 0 Then '������ ����� ����������� ������ ��� �����, �.�. ���� ���� SQRT - �� ��� ������ ����� ����� 1D!!!
        IsCorrectShapeForLabel = False
    End If
'---�������� �� ���������
    If Not ShpTO.Cells("BegTrigger").FormulaU = "" Or Not ShpTO.Cells("EndTrigger").FormulaU = "" Then '� ����������� ��� ������ ������ ���������!!!
        IsCorrectShapeForLabel = False
    End If
'---�������� ������� �������
    If ShpTO.SectionExists(visSectionProp, 0) = False Then
        IsCorrectShapeForLabel = False
    End If
    
End Function

Sub InsertLink(ShpTO As Visio.Shape, ShpFROM As Visio.Shape)
'��������� ���������� ������� ������� ������ ��������� � ������� �������
'---��������� ����������
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX

'---���������� ��������� � ��������� ������ ������� � �������
    '---���������� ��������� � ��������� ������� ������� � �������
    Set mstrConnection = ThisDocument.Masters("���������")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set Cell1 = shpConnection.CellsU("BeginX")
    Set Cell2 = ShpTO.CellsSRC(1, 1, 0)
        Cell1.GlueTo Cell2
    Set Cell1 = shpConnection.CellsU("EndX")
    Set Cell2 = ShpFROM.CellsSRC(1, 1, 0)
        Cell1.GlueTo Cell2
        
        
    '---������ �������� � ConnectionPoints ������ �������
    CellFormula = "PAR(PNT(Sheet." & ShpFROM.ID & "!Connections.ConPoint.X," & _
                          "Sheet." & ShpFROM.ID & "!Connections.ConPoint.Y))"
        shpConnection.CellsU("EndX").FormulaU = CellFormula
        shpConnection.CellsU("EndY").FormulaU = CellFormula
    
    '---������ ��������� ��������� �������� ��� �������
    CellFormula = "IF(Sheet." & shpConnection.ID & "!BeginX<PinX,Width*0,Width*1)"
        ShpFROM.CellsU("Connections.ConPoint.X").FormulaU = CellFormula
    
'---���������� �������� ������ ����������
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
      
   CellFormula = "Sheet." & ShpFROM.ID & "!LineColor"
    shpConnection.Cells("LineColor").FormulaU = CellFormula
   
'---������ ����� �� �������
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpFROM, visSelect
    

    
Exit Sub
EX:
    SaveLog Err, "InsertLabelSquare"
'---������ ����� �� �������
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpFROM, visSelect
'---������� ���������
    shpConnection.Delete
End Sub


Private Function GetPropsList(ByRef DirShpObj As Visio.Shape) As String
'�������� �������� ���� ������� ��������� ������
Dim i As Integer
Dim tempStr As String
    
    tempStr = "����������������"
    If DirShpObj.SectionExists(visSectionProp, 0) = True Then
    '---���� ������ ���� - ��������� ������
        For i = 0 To DirShpObj.RowCount(visSectionProp) - 1
            If DirShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).Result(visNone) = 0 Then
                tempStr = tempStr & ";" & DirShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone)
            End If
        Next i
    End If

GetPropsList = tempStr
End Function

Private Function GetPropsLinks(ByRef DirShpObj As Visio.Shape) As String
'�������� �������� ������ �� �������� ��������� ������
Dim i As Integer
Dim tempStr As String
    
    tempStr = "Prop.Property.Prompt"
    If DirShpObj.SectionExists(visSectionProp, 0) = True Then
    '---���� ������ ���� - ��������� ������
        For i = 0 To DirShpObj.RowCount(visSectionProp) - 1
            If DirShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).Result(visNone) = 0 Then
                Select Case DirShpObj.CellsSRC(visSectionProp, i, visCustPropsType).Result(visNumber)
                    Case Is = 2
                        tempStr = tempStr & Chr(38) & Chr(34) & ";" & Chr(34) & Chr(38) & _
                                    "Round(Sheet." & DirShpObj.ID & "!" & _
                                    DirShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).Name & ",Prop.Round)"
                    Case Is = 5
                        tempStr = tempStr & Chr(38) & Chr(34) & ";" & Chr(34) & Chr(38) & _
                                    "Format(Sheet." & DirShpObj.ID & "!" & _
                                    DirShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).Name & "," & Chr(34) & "HH:mm" & Chr(34) & ")"
                    Case Else
                        tempStr = tempStr & Chr(38) & Chr(34) & ";" & Chr(34) & Chr(38) & _
                                    "Sheet." & DirShpObj.ID & "!" & _
                                    DirShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).Name
                End Select

            End If
        Next i
    End If

GetPropsLinks = tempStr

End Function

Private Sub GetCorrectShape(ByRef shp As Visio.Shape)
Dim corrShp As Visio.Shape
    
    On Error GoTo EX
    Set corrShp = shp.Parent
    '---���� ������������ ������ != ��������
    GetCorrectShape corrShp
    Set shp = corrShp
    
Exit Sub
EX:
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
