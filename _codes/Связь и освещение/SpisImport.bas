Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)
'---��������� ����� ������
ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

'---��������� ������ ������� � �� ���
ModelsListImport (ShpObj.ID)

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
Select Case IndexPers
    Case Is = 58 '������� ������������
        Criteria = "�������"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("������������", "������", "���", Criteria)
    Case Is = 59   '"�������������"
        Criteria = "�������������"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("������������", "������", "���", Criteria)
    Case Is = 23 '������������ ������������
        Criteria = "������������"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("������������", "������", "���", Criteria)
        
End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "ModelsListImport"
End Sub








