Attribute VB_Name = "ObjectProperties"
'----------------------------------������ ��������� ������� �������--------------------------------


Public Sub sP_ChangeObjectProperties(ShpObj As Visio.Shape)
ShpObj.Delete
'��������� ��������� ������� ������� ������
'---��������� ����������
Dim vpVS_DocShape As Visio.Shape

'---���������� ������ ����-����� ���������
Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---��������� ������� �� � ������ User ����-����� ��������� ������ City
If vpVS_DocShape.CellExists("User.City", 0) = False Then
    sp_DocumentRowsAdd
End If

'---��������� ���� �������
PropertiesForm.Show

    
End Sub

Private Sub sp_DocumentRowsAdd()
'��������� �������� ����� ��� ������� ������� ������
'---��������� ����������
Dim vpVS_DocShape As Visio.Shape

    On erro GoTo EX
'---���������� ������ ����-����� ���������
    Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---��������� ������� ������ User � � ������ � ���������� �������
    If vpVS_DocShape.SectionExists(visSectionUser, 0) = False Then
        vpVS_DocShape.AddSection (visSectionUser)
    End If

'---��������� ����� ������
    '---������ "���������� �����"
    vpVS_DocShape.AddNamedRow visSectionUser, "City", visTagDefault
    vpVS_DocShape.Cells("User.City").FormulaU = """�������� ����������� ������"""
    vpVS_DocShape.Cells("User.City.Prompt").FormulaU = """�������� ����������� ������"""
    '---������ "�����"
    vpVS_DocShape.AddNamedRow visSectionUser, "Adress", visTagDefault
    vpVS_DocShape.Cells("User.Adress").FormulaU = """����� ������� ������"""
    vpVS_DocShape.Cells("User.Adress.Prompt").FormulaU = """����� ������� ������"""
    '---������ "������ ������"
    vpVS_DocShape.AddNamedRow visSectionUser, "Object", visTagDefault
    vpVS_DocShape.Cells("User.Object").FormulaU = """������ ������"""
    vpVS_DocShape.Cells("User.Object.Prompt").FormulaU = """������ ������"""
    '---������ "������� ������������� �������"
    vpVS_DocShape.AddNamedRow visSectionUser, "FireRating", visTagDefault
    vpVS_DocShape.Cells("User.FireRating").FormulaU = """3"""
    vpVS_DocShape.Cells("User.FireRating.Prompt").FormulaU = """������� �������������"""

Exit Sub
EX:
    SaveLog Err, "sp_DocumentRowsAdd"
End Sub


'Public Function PPP(ShpObj As Visio.Shape) As Boolean
'    ShpObj.Cells("User.Row_1") = 111
'End Function
