Attribute VB_Name = "ValueImportModule"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub ProductionImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ���������� � ����������� � ����� �������� � �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ���� ����� ��� �������� ���� ����� ��� ������� ������
    Select Case IndexPers
        Case Is = 50 '�������� �������
            Criteria = "[��� ��������] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' And " & _
                "[������� ��������] = " & shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString) & " And " & _
                "[����� � ����] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
            shp.Cells("Prop.Production").FormulaForceU = "Guard(" & ValueImportSng("����������������", "����������", Criteria) & ")"
    
    End Select

Set shp = Nothing
Exit Sub
EX:
    SaveLog Err, "ProductionImport"
End Sub


