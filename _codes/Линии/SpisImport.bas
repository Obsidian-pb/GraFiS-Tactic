Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)

'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---��������� ����� ������
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
    
        '---��������� ��� ����� ������ ����������� ��������� � ��������� ��������� ������
        Select Case ShpObj.Cells("User.IndexPers")
            Case Is = 34 '������� ������ �����
                ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("��������������������", "������ ������")
            Case Is = 36 '�������� ������� �����
                ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("���������������������", "������ ������")
            Case Is = 35 '������ ������ �����
                ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("�������������������", "������ ������")
                ShpObj.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("����������������", "����������������")
            Case Is = 37 '������ �������� �����
                ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("��������������������", "������ ������")
                ShpObj.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("����������������", "����������������")
            Case Is = 39 '������� ������� �������� �����
                ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("����������������������", "������ ������")
            Case Is = 76 '������ ������� �����
                ShpObj.Cells("Prop.Gas.Format").FormulaU = ListImport("������� �������", "������")
            Case Is = 77 '������ ���������� �����
                ShpObj.Cells("Prop.Powder.Format").FormulaU = ListImport("�������", "�����")
            Case Is = 40 '�������������
                ShpObj.Cells("Prop.WEType.Format").FormulaU = ListImport("��������������", "������")
            Case Is = 41 '������������� - ���� ������ �� ����������!!!
                
            Case Is = 88 '����� �����������
                ShpObj.Cells("Prop.Model.Format").FormulaU = ListImport("����� �����������", "������")
            Case Is = 72 '������� ��������
                ShpObj.Cells("Prop.ColPressure.Format").FormulaU = ListImportInt("�������", "�����")
        End Select
        
        '---��������� ������ �� ������� ����� ��������
        ShpObj.Cells("Prop.SetTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If



    
'---������� ������� ���������� ��� ������
    ConnectionsRefresh ShpObj
    
    
On Error Resume Next '�� ������ ��� ������� ���������� ������
Application.DoCmd (1312)

End Sub

Public Sub UnitsListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (������ �������������)
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

End Sub

Public Sub WFListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (����� �����������)
'Dim IndexPers As Integer

    '---��������� ����� ������
'    ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("��������������������", "������ ������")
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
    ShpObj.Cells("Prop.WFType.Format").FormulaU = ListImport("����� �����������", "������")
    
    
On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

End Sub

Public Sub StvolModelsListImport(ShpIndex As Long)
'��������� ������� ������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 34 '������� ������ �����
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("��������������������", "������ ������")
        Case Is = 36 '�������� ������� �����
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("���������������������", "������ ������")
        Case Is = 35 '������ �����
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("�������������������", "������ ������")
        Case Is = 37 '������ �������� �����
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("��������������������", "������ ������")
        Case Is = 39 '�������� ������� ������� �����
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("����������������������", "������ ������")
            
            
            
    End Select
    
    '---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.StvolType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StvolType").FormulaU = "INDEX(0,Prop.StvolType.Format)" '!!! ��� ����� ���� ������!!!!
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolModelsListImport"
End Sub

Public Sub WEModelsListImport(ShpIndex As Long)
'��������� ������� ������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
    shp.Cells("Prop.WEType.Format").FormulaU = ListImport("��������", "������")
        
'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.WEType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.WEType").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "WEModelsListImport"
End Sub

Public Sub WFModelsListImport(ShpIndex As Long)
'��������� ������� ������� ����������� �����
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
    shp.Cells("Prop.WFType.Format").FormulaU = ListImport("����� �����������", "������")
        
'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.WFType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.WFType").FormulaU = "INDEX(0,Prop.WFType.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "WFModelsListImport"
End Sub

Public Sub StvolFoamCreatorListImport(ShpIndex As Long)
'��������� ������� ������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
Select Case IndexPers
    Case Is = 35 '������ ������ �����
        shp.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("����������������", "����������������")
    Case Is = 37 '������ �������� �����
        shp.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("����������������", "����������������")
        
        
        
        
        
End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.StvolType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StvolType").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolFoamCreatorListImport"
End Sub




'------------------------���� ��������� �������------------------------------
Public Sub StvolVariantsListImport(ShpIndex As Long)
'��������� ������� ��������� �������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ������ ������� ��� ������� ������
Select Case IndexPers
    Case Is = 34
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("��������������������", "������� ������", Criteria)
    Case Is = 36
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("���������������������", "������� ������", Criteria)
    Case Is = 35 ' ������ ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("�������������������", "������� ������", Criteria)
    Case Is = 37 ' ������ �������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("��������������������", "������� ������", Criteria)
    Case Is = 39 '�������� ������� ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("����������������������", "������� ������", Criteria)
        
        
        
        
End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.Variant").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Variant").FormulaU = "INDEX(0,Prop.Variant.Format)"
    End If
    
    Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolVariantsListImport"
End Sub


Public Sub StvolStreamTypesListImport(ShpIndex As Long)
'��������� ������� ������ ��������� ����� ��� ������� ������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ��������� ����� ������� ��� ������� ������
Select Case IndexPers
    Case Is = 34
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("��������������������", "��� �����", Criteria)
    Case Is = 36
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("���������������������", "��� �����", Criteria)
    Case Is = 35
        Set shp = Nothing
        Exit Sub
    Case Is = 37
        Set shp = Nothing
        Exit Sub
    Case Is = 39 '������� ������� �������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("����������������������", "��� �����", Criteria)
        
        
        
        
End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.StreamType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StreamType").FormulaU = "INDEX(0,Prop.StreamType.Format)"
    End If
    
    Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolStreamTypesListImport"
End Sub



Public Sub StvolHeadListImport(ShpIndex As Long)
'��������� ������� ������ ��������� ������� ��� ������� ���� ����� � ������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---��������� ��������� ��������� �������������� ������ ��������� ������� ��� ������� ������
Select Case IndexPers
    Case Is = 34
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("��������������������", "�����", Criteria)
    Case Is = 36
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("���������������������", "�����", Criteria)
    Case Is = 35 '������ ������ �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("�������������������", "�����", Criteria)
    Case Is = 37 '������ �������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("��������������������", "�����", Criteria)
    Case Is = 39 '������� �������� ������� �����
        Criteria = "[������ ������] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[������� ������] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[��� �����] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("����������������������", "�����", Criteria)
    Case Is = 40 '�������������
        Criteria = "[������] = '" & shp.Cells("Prop.WEType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Pressure.Format").FormulaU = ListImport2("��������", "����� �� �����", Criteria)
        
        
        
End Select

'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.Head").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Head").FormulaU = "INDEX(0,Prop.Head.Format)"
    End If
    
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "StvolHeadListImport"
End Sub



