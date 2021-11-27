Attribute VB_Name = "Functions"
Sub JPGExportAll(ShpObj As Visio.Shape)
    ShpObj.Delete
    ExportJPG.Show
End Sub




Sub SetAspect(ShpObj As Visio.Shape)
'��������� ������������� ����� �������� ������� ��� ������ ��������
    SetAspect_P
    ShpObj.Delete
End Sub


Public Sub FixZIndex(ShpObj As Visio.Shape)
'����� ���������� ��������� ����� ������ ������������ ���� � �������
    FixZIndex_P
    ShpObj.Delete
End Sub


'-------------------------����� ��� ������ �� ������--------------------------------
Sub JPGExportAll_P()
    ExportJPG.Show
End Sub

Sub SetAspect_P()
'��������� ������������� ����� �������� ������� ��� ������ ��������
Dim vO_Sheet As Visio.Shape
Dim vs_Aspect As Single

'---���������� ���������� ���������
    Dim UndoScopeID As Long

Set vO_Sheet = Application.ActiveWindow.Shape

On Error GoTo ex

'---�������� ������ �� �� �������� ������ GFS_Aspect
    If vO_Sheet.CellExists("User.GFS_Aspect", 0) = False Then
    '---���� ���, �� ������� �� ��������� 1
        If vO_Sheet.SectionExists(visSectionUser, 0) = False Then '��������� ������� �� ������, ������� - �������
            vO_Sheet.AddSection visSectionUser
        End If
            vO_Sheet.AddNamedRow visSectionUser, "GFS_Aspect", visTagDefault
            vO_Sheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---���������� ��� ��������
    vs_Aspect = _
    CSng(InputBox("�������� �������� ������� �� ������ �������. ������ ��������� ������ �������������� ��������������� ��� ���� ����� ������, ��� ����� ���� ������ ��� ������ � ������ � ������������ ���������.", _
        "������ - ��������� �������", vO_Sheet.Cells("User.GFS_Aspect").Result(visNumber)))
    
'---��������� ������������ �������� �������
    If vs_Aspect <= 0 Or vs_Aspect > 100 Then
        GoTo ex
    End If

'---������������� ����� �������� �������
    vO_Sheet.Cells("User.GFS_Aspect").Formula = vs_Aspect

Set vO_Sheet = Nothing
Application.EndUndoScope UndoScopeID, True

Exit Sub

ex:
MsgBox "�������� ���� �������� �� ����� ���� ����������� � �������� �������! ��������� ��������� �� �� ��� �������! � �������� �������� ����� ���� ������������ ������ ����� �� 0,1 �� 100!", vbCritical, ThisDocument.Name
Set vO_Sheet = Nothing
Application.EndUndoScope UndoScopeID, True

End Sub


Public Sub FixZIndex_P()
'����� ���������� ��������� ����� ������ ������������ ���� � �������
Dim vsoSelection As Visio.Selection
    
    On Error GoTo ex
    
    '---���������� ������ //�������;���;�������� �����;�������������;����
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "�������;���;�������� �����;�������������;����")
    Application.ActiveWindow.Selection = vsoSelection
    
    Application.ActiveWindow.Selection.BringToFront
    
    ActiveWindow.DeselectAll
    
    '---���������� ������ //����;������� �������
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "����;������� �������;����;���������� ���")
    Application.ActiveWindow.Selection = vsoSelection
    
    Application.ActiveWindow.Selection.BringToFront
    
    ActiveWindow.DeselectAll
    
    
Exit Sub
ex:
    
End Sub


Public Sub ShapesCountShow()
'����� ������ ���������� ����� � �������
Dim vO_ShpItm As Visio.Shape
Dim x, y As Double
    
    On Error GoTo ex
    
    '���������� �� ����� ������ �������� �������
    MsgBox "���������� ����� � ���������: " & Application.ActiveWindow.Selection.Count, , "������"
    
Exit Sub
ex:
End Sub
