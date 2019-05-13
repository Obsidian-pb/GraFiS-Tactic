Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

    '---��������� ����� ������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

    '---��������� ��� ����� ������ ����������� ��������� � ��������� ��������� ������
    Select Case ShpObj.Cells("User.IndexPers")
        Case Is = 100 '������� ������ �����
            ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("�_������", "�������� ������")
            
    End Select
    
    
On Error Resume Next '�� ������ ��� ������� ���������� ������
Application.DoCmd (1312)

End Sub

Public Sub InLineListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

    '---��������� ����� ������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

On Error Resume Next '�� ������ ��� ������� ���������� ������
Application.DoCmd (1312)

End Sub

Public Sub DropNewShape(ShpObj As Visio.Shape)
'��������� ����� ������ ����� �����
    If IsFirstDrop(ShpObj) Then
        '---��������� ������ �� ������� ����� ��������
        If ShpObj.Cells("User.IndexPers").Result(visUnitsNone) = 102 Then                           '����
            ShpObj.Cells("Prop.BreakeupTime").Formula = _
                Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
        ElseIf ShpObj.Cells("User.IndexPers").Result(visUnitsNone) = 103 Then                       '������
            ShpObj.Cells("Prop.SetTime").Formula = _
                Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
        End If
        
        
    End If

End Sub

'------------------------���� ��������� �������------------------------------
Public Sub HoseDiametersListImport(ShpIndex As Long)
'��������� ������� ��������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
Select Case indexPers
    Case Is = 100
        Criteria = "[�������� ������] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.HoseDiameter.Format").FormulaU = ListImport2("�_������", "������� �������", Criteria)

End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
If shp.Cells("Prop.HoseDiameter").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.HoseDiameter").FormulaU = "INDEX(0,Prop.HoseDiameter.Format)"
End If

Set shp = Nothing

End Sub


