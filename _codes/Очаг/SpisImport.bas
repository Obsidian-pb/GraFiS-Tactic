Attribute VB_Name = "SpisImport"
'------------------------������ ��� �������� ������� �������-------------------
'------------------------���� ����������� �������------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'��������� ���������� ������ ������ (���� �������)
'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---��������� ������ ���������
            ShpObj.Cells("Prop.FireCategorie.Format").FormulaU = ListImport("�_�������������", "���������")
        
        '---��������� ������ �������� � ������������ �� ��������� ���������
            DescriptionsListImport (ShpObj.ID)
        '---��������� ��������� ��������� �������� �������� ������ ��� ������� ��������
            GetFactorsByDescription (ShpObj.ID)
        
        '---��������� ������ �� ������� ����� ��������
        ShpObj.Cells("Prop.SquareTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If

On Error Resume Next '�� ������ ��� ������� ���������� ������
If VfB_NotShowPropertiesWindow = False Then Application.DoCmd (1312) '� ������ ���� ����� ���� �������, ���������� ����

End Sub

Public Sub SetRushTime(ShpObj As Visio.Shape)
'��������� ������������� ����� ��������� ������� ��������
'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---��������� ������ �� ������� ����� ��������
        ShpObj.Cells("Prop.RushTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If


End Sub

Public Sub DescriptionsListImport(ShpIndex As Long)
'��������� ������� ������ ��������
'---��������� ����������
Dim shp As Visio.Shape
Dim Criteria As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---������������� �������� ���������
    Criteria = shp.Cells("Prop.FireCategorie").ResultStr(Visio.visNone)

'---��������� ��������� ��������� �������������� ������ ��������� ��� ������� ������
        If shp.Cells("Prop.IntenseShowType").ResultStr(Visio.visNone) = "�� ���������" Then
            shp.Cells("Prop.FireDescription.Format").FormulaU = ListImport2("�_�������������", "��������", "���������", Criteria)
        End If

'---� ������, ���� �������� ���� ��� ������� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
    If shp.Cells("Prop.FireDescription.Format").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.FireDescription").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.FireDescription").FormulaU = "INDEX(0,Prop.FireDescription.Format)"
    End If

Set shp = Nothing
End Sub











Private Sub ToZeroListIndex(cell As String, ShpIndex As Long) '!!!�������� �� ������������ � ����� � ����������� ������������
'---� ������, ���� �������� ���� ��� ������ ������ ����� "", ��������� ����� � ������ �� 0-� ���������.
Dim CellName As String, CellContent As String

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    
'---���������� �������� ����� ������� ����� �������� ��������
    CellName = "Prop." & cell
    CellContent = "INDEX(0,Prop." & cell & ".Format)"
    If shp.Cells(CellName).ResultStr(Visio.visNone) = "" Then
        shp.Cells(CellName).FormulaU = CellContent
    End If
End Sub




