Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

'---��������� ����� ������
    ShpObj.Cells("Prop.PipeType.Format").FormulaU = ListImport("��� ����", "��� ��������")
    
'---��������� ������ ������� � �� ���
    '---��������� ��������� ��������� ������ ���������
    DiametersListImport (ShpObj.ID)
    '---��������� ��������� ��������� ������ �������
    PressuresListImport (ShpObj.ID)
    '---��������� ��������� ��������� ����������
    ProductionImport (ShpObj.ID)

On Error Resume Next '�� ������ ��� ������� ���������� ������
Application.DoCmd (1312)

End Sub

'------------------------���� ��������� �������------------------------------

Public Sub DiametersListImport(ShpIndex As Long)
'��������� ������� ������� ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
    Select Case IndexPers
        Case Is = 50 '�������� �������
            Criteria = "[��� ��������] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' "
            shp.Cells("Prop.PipeDiameter.Format").FormulaU = ListImportNum("����������������", "������� ��������", Criteria)
    
    End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
'    If shp.Cells("Prop.PipeDiameter").ResultStr(Visio.visNone) = "" Then
'        shp.Cells("Prop.PipeDiameter").FormulaU = "INDEX(0,Prop.PipeDiameter.Format)"
'    End If
    
    Set shp = Nothing

Exit Sub
EX:
    SaveLog Err, "DiametersListImport", CStr(ShpIndex)
End Sub


Public Sub PressuresListImport(ShpIndex As Long)
'��������� ������� ������ ��������� ������� ��� �������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ��������� ����� ������� ��� ������� ������
    Select Case IndexPers
        Case Is = 50
            Criteria = "[��� ��������] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' And " & _
                "[������� ��������] = " & shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.Pressure.Format").FormulaU = ListImportNum("����������������", "����� � ����", Criteria)
    End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", �������� ���������������� ����
    If shp.Cells("Prop.Pressure").ResultStr(Visio.visNone) = "" Then
        'shp.Cells("Prop.Pressure").FormulaU = "INDEX(0,Prop.Pressure.Format)"
        shp.Cells("Prop.ShowDirectProduction").FormulaU = "INDEX(1,Prop.ShowDirectProduction.Format)"
    End If

    Set shp = Nothing

Exit Sub
EX:
    SaveLog Err, "PressuresListImport", ShpIndex
End Sub


