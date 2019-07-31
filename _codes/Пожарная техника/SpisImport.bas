Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------


Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        If Not IsShapeLinked(ShpObj) Then
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
    End If

'---������� �������� ������������ �������� �����
    ConnectionsRefresh ShpObj

On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

Exit Sub
EX:
    Application.EventsEnabled = True
    SaveLog Err, "BaseListsRefresh", "�������� �������"
End Sub
Public Sub BaseListsRefresh2(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������) ������ ��� ����� ������
'---��������� ����� ������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
    
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

    On Error GoTo Tail

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
    Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 1
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_������������", "������", "�����", Criteria)
        Case Is = 2
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 3
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 4
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 5
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 6
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 7
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 8
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 9
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 10
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 11
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 12
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 13
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_����", "������", "�����", Criteria)
        Case Is = 14
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 15
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_�����", "������", "�����", Criteria)
        Case Is = 16
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 17
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 18
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 19
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 20
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_��", "������", "�����", Criteria)
        Case Is = 161 '���
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
        Case Is = 162 '����
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_����", "������", "�����", Criteria)
        Case Is = 163 '���
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("�_���", "������", "�����", Criteria)
            
            
    End Select
    
'---� ������, ���� �������� ���� ��� ������� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.Model.Format").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
    
    Set shp = Nothing
    
Exit Sub
Tail:
    SaveLog Err, "ModelsListImport"
End Sub



