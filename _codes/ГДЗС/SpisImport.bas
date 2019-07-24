Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        If IsFirstDrop(ShpObj) Then
            '---��������� ����� ������
            ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
            'ShpObj.Cells("Prop.AirDevice.Format").FormulaU = ListImport("����", "������")
            
            '---��������� ������ ������� � �� ���
            '---��������� ��� ����� ������ ����������� ��������� � ��������� ��������� ������ (������ ��� �������)
            Select Case ShpObj.Cells("User.IndexPers")
                Case Is = 46 '����
                    AirDevicesListImport (ShpObj.ID)
                    GetTTH (ShpObj.ID)
                Case Is = 90 '����
                    AirDevicesListImport (ShpObj.ID)
                    GetTTH (ShpObj.ID)
            End Select
            
            '---��������� ������ �� ������� ����� ��������                                                                                         '��� ���� ���������
                ShpObj.Cells("Prop.FormingTime").Formula = _
                    Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    
        End If
    End If

On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

End Sub

Public Sub FogRMKBaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������) ��� ���������
'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---��������� ����� ������
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
        
        '---��������� ������ ������� � �� ���
        '---��������� ������ ������� � �� ���
        FogRMKListImport (ShpObj.ID)
        GetTTH (ShpObj.ID)
        
        '---��������� ������ �� ������� ����� ��������
            ShpObj.Cells("Prop.SetTime").Formula = _
                Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If
    




On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

End Sub

'------------------------���� ��������� �������------------------------------
Public Sub AirDevicesListImport(ShpIndex As Long)
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
        Case Is = 46
            shp.Cells("Prop.AirDevice.Format").FormulaU = ListImport("����", "������")
        Case Is = 90
            shp.Cells("Prop.AirDevice.Format").FormulaU = ListImport("����", "������")
    End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.AirDevice").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.AirDevice").FormulaU = "INDEX(0,Prop.AirDevice.Format)"
    End If
    
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    SaveLog Err, "AirDevicesListImport", CStr(ShpIndex)
End Sub


Public Sub FogRMKListImport(ShpIndex As Long)
'��������� ������� ������� ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
    shp.Cells("Prop.FogRMK.Format").FormulaU = ListImport("��������", "������")

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.FogRMK").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.FogRMK").FormulaU = "INDEX(0,Prop.FogRMK.Format)"
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    SaveLog Err, "FogRMKListImport", CStr(ShpIndex)
End Sub
