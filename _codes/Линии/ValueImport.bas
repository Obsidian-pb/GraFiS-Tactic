Attribute VB_Name = "ValueImport"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub HoseResistanceValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ������������� � ����������� � ���������� � ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���������� ������������� ��������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.HoseResistance").FormulaU = str(ValueImportSng("�_������", "�������������", Criteria))
    
    End Select

Set shp = Nothing

End Sub

Public Sub HoseMaxFlowValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ���������� ����������� � ����������� � ���������� � ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.FlowS").FormulaU = str(ValueImportSng("�_������", "������", Criteria))
    
    End Select

Set shp = Nothing

End Sub

Public Sub HoseWeightValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ����� �_������ � ����������� � ���������� � ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.HoseWeight").FormulaU = str(ValueImportSng("�_������", "�����", Criteria))
    
    End Select

Set shp = Nothing

End Sub
