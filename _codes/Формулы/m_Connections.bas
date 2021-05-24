Attribute VB_Name = "m_Connections"
Option Explicit



Public Sub TurnIntoFormulaConnection(ByRef Connects As IVConnects)
'��������� ��������� � ��������� ������ (��� ����������)
Dim cnct As Visio.Connect
Dim shp As Visio.Shape
Dim rowI As Integer
    
    On Error GoTo EndSub
    
    '���������� ������ ���������
    Set shp = Connects(1).FromSheet
    
    '---��������� �� �������� �� ������ ������� ���������� ���
    If IsGFSShapeWithIP(shp, 501, True) Then Exit Sub
    
    '---���������, ����� �� ������ ��� ����� ����������
    If shp.Connects.Count <> 2 Then Exit Sub

    '---��������� �������� �� ������ ������
    If shp.AreaIU > 0 Then Exit Sub
    
    '--��������� ��� �� ����������� ����� �������� �������� ������
    If IsGFSShapeWithIP(shp.Connects(1).ToSheet, 500, True) And IsGFSShapeWithIP(shp.Connects(2).ToSheet, 500, True) Then
        '---�������� ��������� ���������
        f_LinkToCell2.showForm shp.Connects(1).ToSheet, shp.Connects(2).ToSheet, shp
        
        '��������� ���������� ��������
        shp.AddNamedRow visSectionUser, "IndexPers", visTagDefault
        SetCellVal shp, "User.IndexPers", "501"
    End If
    
Exit Sub
EndSub:
    SaveLog Err, "TurnIntoFormulaConnection"
End Sub
