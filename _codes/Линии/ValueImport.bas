Attribute VB_Name = "ValueImport"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub HoseResistanceValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ������������� � ����������� � ���������� � ���������
'---��������� ����������
Dim Shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set Shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = Shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���������� ������������� ��������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & Shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & Shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
'            Debug.Print CStr(ValueImportSng("�_������", "�������������", Criteria))
            Shp.Cells("Prop.HoseResistance").FormulaU = """" & CStr(ValueImportSng("�_������", "�������������", Criteria)) & """"
    End Select

Set Shp = Nothing

End Sub

Public Sub HoseMaxFlowValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ���������� ����������� � ����������� � ���������� � ���������
'---��������� ����������
Dim Shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set Shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = Shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & Shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & Shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            Shp.Cells("Prop.FlowS").FormulaU = str(ValueImportSng("�_������", "������", Criteria))
    
    End Select

Set Shp = Nothing

End Sub

Public Sub HoseWeightValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ����� �_������ � ����������� � ���������� � ���������
'---��������� ����������
Dim Shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set Shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = Shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case indexPers
        Case Is = 100
            Criteria = "[�������� ������] = '" & Shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[������� �������] = " & Shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            Shp.Cells("Prop.HoseWeight").FormulaU = str(ValueImportSng("�_������", "�����", Criteria))
    
    End Select

Set Shp = Nothing

End Sub
