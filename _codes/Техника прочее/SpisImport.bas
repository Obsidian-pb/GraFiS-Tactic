Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)
'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---��������� ����� ������
        ShpObj.Cells("Prop.Set.Format").FormulaU = ListImport("������", "�����")
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
        
        '---��������� ������ ������� � �� ���
        ModelsListImport (ShpObj.ID)
        GetTTH (ShpObj.ID)
        
        '---��������� ������ �� ������� ����� ��������
        ShpObj.Cells("Prop.ArrivalTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If


On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

End Sub

'------------------------���� ��������� �������------------------------------
Public Sub ModelsListImport(ShpIndex As Long)
'��������� ������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 73 ' ������ �� ���������� ����
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[���] = '������ �� ���������� ����'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_���������� ������", "������", Criteria)
        Case Is = 74 ' �����
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[���] = '����'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_���������� ������", "������", Criteria)
        Case Is = 30 '�������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[�����] = '����'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_����", "������", Criteria)
        Case Is = 31 '�����
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[�����] = '����'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_����", "������", Criteria)
        Case Is = 24 '������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_������", "���������", Criteria)
        Case Is = 28 '���������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_���������", "������", Criteria)
        Case Is = 25 '��������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[���] = '�������'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_��������", "������", Criteria)
        Case Is = 26 '��������-�������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[���] = '�������'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_��������", "������", Criteria)
        Case Is = 27 '���������
            Criteria = "[�����] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("�_���������", "������", Criteria)
    
    End Select

'---� ������, ���� �������� ���� ��� ������� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.Model.Format").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ModelsListImport"
End Sub








