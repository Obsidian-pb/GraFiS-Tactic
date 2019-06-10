Attribute VB_Name = "ValueImportModule"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub StvolStreamValueImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ��� ����� � ����������� � ����� �����
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
    Case Is = 34
        Criteria = "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("�����", "��� �����", Criteria)
    Case Is = 36
        Criteria = "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("�����", "��� �����", Criteria)
    Case Is = 39
        Criteria = "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("�����", "��� �����", Criteria)
        
        
        
        
End Select

Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolStreamValueImport"
End Sub

Public Sub StvolDiameterInImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ��� ����� � ����������� � ����� �����
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
    Case Is = 34
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("�������������", "�������� ������", Criteria)
    Case Is = 36
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("�������������", "�������� ������", Criteria)
    Case Is = 35 '������ ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("�������������", "�������� ������", Criteria)
    Case Is = 37 '������ �������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("�������������", "�������� ������", Criteria)
    Case Is = 39 '�������� ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("�������������", "�������� ������", Criteria)
        
        
End Select

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolDiameterInImport"
End Sub

Public Sub StvolProductionImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ����� � ����������� � ������� �������
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
    Case Is = 34   ' ������ ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[�����] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("��������������������", "������", Criteria))
    Case Is = 36   ' �������� ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[�����] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("���������������������", "������", Criteria))
    Case Is = 35   ' ������ ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[�����] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("�������������������", "������", Criteria))
    Case Is = 37   ' �������� ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[�����] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("��������������������", "������", Criteria))
    Case Is = 39   ' �������� ������� ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[�����] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("����������������������", "������", Criteria))
    Case Is = 40   ' �������������
        '���� - �������� �� ������������ ������� � �������
'        Criteria = "[������] = '" & shp.Cells("Prop.WEType").ResultStr(visUnitsString) & "'  And " & _
'            "[����� �� �����] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
'        shp.Cells("Prop.PodOut").FormulaU = Str(ValueImportSng("��������", "������������������", Criteria))
'        shp.Cells("Prop.PressureOut").FormulaU = Str(ValueImportSng("��������", "������������������", Criteria))
    Case Is = 88   ' ����������� �����
        Criteria = "[������] = '" & shp.Cells("Prop.WFType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.PodIn").FormulaU = str(ValueImportSng("����� �����������", "������������������", Criteria))
        
        
        
End Select


Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolProductionImport"
End Sub

Public Sub StvolWFLinkImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ WFLink (������ �� ��������� ����� wiki-fire.org)
'� ������������ � ������� ������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
    
'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
'    IndexPers = shp.Cells("User.IndexPers")
    
'---��������� ������ � �� � �������� �������� ������ �� wiki-fire.org
    Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "'"
    shp.Cells("Prop.WFLink").FormulaU = ValueImportStr("�������������", "������ WF", Criteria)
   
Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolWFLinkImport"
End Sub

Public Sub StvolWFLinkFree(ShpIndex As Long)
'����� ������������� � �������� ������ ������ ��������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
    
'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    
'---������������� ������ ������, ����� ������� ��������� ������ ������������ ������ ������ ��-���������
    shp.Cells("Prop.WFLink").FormulaU = ""
   
Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolWFLinkImport"
End Sub

Public Sub StvolRFImport(ShpIndex As Long)
'��������� ������������ � ������������� ������ ��������� � ����������� � ������� ������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ��������� ��� �������� ������ ������ ��� ������� ������
Select Case IndexPers

    Case Is = 35   ' ������ ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.FoamRF").FormulaU = str(ValueImportSng("�������������������", "���������", Criteria))
    Case Is = 37   ' �������� ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.FoamRF").FormulaU = str(ValueImportSng("��������������������", "���������", Criteria))
        
        
        
        
End Select

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolRFImport"
End Sub


Public Sub ColFlowMaxImport(ShpIndex As Long)
'��������� ������������� �������� ������ ������������ � ����������� � �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
'    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� ��������� ��� �������� ������ ������ ��� ������� ������
'---���������� �������� ������ - ����� ����� ��������
    Criteria = "[�����] = " & shp.Cells("Prop.ColPressure").ResultStr(visUnitsString)
'---��������� ����� �������� � ������� � � ������������ � ���� �������� �������� �� ��
    If shp.Cells("Prop.Patr").ResultStr(visUnitsString) = "77" Then
        shp.Cells("Prop.FlowMax").FormulaU = str(ValueImportSngStr("�������", "������ 77", Criteria))
    ElseIf shp.Cells("Prop.Patr").ResultStr(visUnitsString) = "66" Then
        shp.Cells("Prop.FlowMax").FormulaU = str(ValueImportSngStr("�������", "������ 66", Criteria))
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ColFlowMaxImport"
End Sub

Public Sub FoamDiffKoeffImport(ShpObj As Visio.Shape)
'��������� ������������� �������� ������ ����������� �������� ������ ��� �������� ������� (� ��������������)
'---��������� ����������
Dim Criteria As String
Dim foamPercentage As String
Dim vstDiameter As Integer

    On Error GoTo EX

'---��������� ��������� ��������� ������������ ��� �������� ���������� ������� (�������������)
    foamPercentage = ShpObj.Cells("Prop.FoamPercentage").ResultStr(visUnitsString)
    vstDiameter = ShpObj.Cells("Prop.FoamInDiameter").Result(visNumber)
    
'---���������� �������� ������ - ����� ����� ��������
    Criteria = "[������������ �������] = " & foamPercentage
'---��������� ������ �������� ������ �������
    If vstDiameter = "10" Then
        ShpObj.Cells("User.DiffKoeff").Formula = str(ValueImportSngStr("�������������������", "����������� �������� 10", Criteria))
    End If
    If vstDiameter = "25" Then
        ShpObj.Cells("User.DiffKoeff").Formula = str(ValueImportSngStr("�������������������", "����������� �������� 25", Criteria))
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "FoamDiffKoeffImport"
End Sub

Public Sub NozzleProvKoeffImport(ShpObj As Visio.Shape)
'��������� ������������� �������� ������ ������������ ������� - User.ProvKoeff
'---��������� ����������
Dim Criteria As String
Dim nozzleDiameter As String

    On Error GoTo EX

'---��������� ��������� ��������� ������������ ��� �������� ���������� ������� (�������)
    nozzleDiameter = ShpObj.Cells("User.NozzleDiameter").ResultStr(visUnitsString)
    
'---���������� �������� ������ - ����� ����� ��������
    Criteria = "[������� �������] = " & Replace(nozzleDiameter, ",", ".")
'---����������� ��������
    ShpObj.Cells("User.ProvKoeff").Formula = CStr(ValueImportSngStr("�����������������", "������������", Criteria))

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "NozzleProvKoeffImport"
End Sub
