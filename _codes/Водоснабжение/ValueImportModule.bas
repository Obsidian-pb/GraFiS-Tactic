Attribute VB_Name = "ValueImportModule"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub ProductionImport(shp As Visio.Shape)
'��������� ������������ � ������������� ������ ���������� � ����������� � ����� �������� � �������
'---��������� ����������
'Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---��������� � ����� ������ ������ ��������� ������ ������
'    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case indexPers
        Case Is = 50 '�������� �������
            Criteria = "[��� ��������] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' And " & _
                "[������� ��������] = " & shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString) & " And " & _
                "[����� � ����] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
            shp.Cells("Prop.Production").FormulaForceU = "Guard(" & ValueImportSng("����������������", "����������", Criteria) & ")"
    
    End Select

'Set shp = Nothing
Exit Sub
EX:
    SaveLog Err, "ProductionImport"
End Sub


