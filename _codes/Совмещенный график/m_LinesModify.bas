Attribute VB_Name = "m_LinesModify"
Option Explicit

'-----------������ ������������ �������� � ����������� � �������� ������� (������ ��� ������ �������� �� �������)
'-----------------------------------------------------------------------------------------------------------------

'--------------------------------������ ������� �������-----------------------------------------------------------
Public Sub GetFireSquareDataFromAnalize(ByRef shp As Visio.Shape)
'����������� ����� ������� ������� ������� � ����������� � ��������
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---�������� �� ������������ ����� �������� ��� �������
    '���������� ����� ��� ������
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  '���� ��� ����� ������ - ������� �� ��������
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---�������� ������ ������ �� �����������
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---���� � ��������� �������� ��� �� ����� ������ (�.�. ��� �� �����) - ������� �� �����
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "�� ��������� �������� ��� ����� �������� �������! ��������� ������ ������ ����������!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---���������, ������� �� ������ �����
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "������ ����� �����������! ��������� ������ ������ ����������!", vbCritical
        Exit Sub
    End If
    
'---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    
    '---�������� ������ ������
    If shp.Cells("User.IndexPers") = 123 Then
        '���� ������ - ������ ������� ������
        GraphAnalizer.PS_GetFireSquares DataArray
    ElseIf shp.Cells("User.IndexPers") = 124 Then
        '���� ������ - ������ ������� �������
        GraphAnalizer.PS_GetExtSquares DataArray
    End If
    
'---�������� ���������� ������ ��������� ���������� ����� �������
    ps_FireGraphicBuild shp, DataArray
    
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetFireSquareDataFromAnalize"
End Sub

Public Sub GetFireSquareDataFromTable(ByRef shp As Visio.Shape)
'����������� ����� ������� ������� ������� � ����������� � �������� ������
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---���� � ����� ����� Cancel, �� ������� �� �����
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    '---�������� �� ����� ������ ������ ��� ���������� ������� �������� � �������� ������
    DataForm.PS_GetMainArray MainArray
    '---������������� ������
    ps_FireGraphicBuild shp, MainArray
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetFireSquareDataFromTable"
End Sub


Private Sub ps_FireGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray())
'����� ������ ����� ������ ������� ������� (�������)
Dim i As Integer

On Error GoTo EX

    '---������������� �������� ��� ������ (���������) �����
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(MainArray(1, 0)) & "/User.FireMax)*Height"
    
    '---������������� �������� ��� ����� ����� (�������� ��)
    For i = 1 To UBound(MainArray, 2)
        PS_FireGraph_AddKnot shp
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(MainArray(1, i)) & "/User.FireMax)*Height"
    Next i
    
Exit Sub
EX:
    MsgBox "������ ���� ������� �� ��������� ����� �����������!", vbCritical
    shp.Delete
End Sub


'--------------------------------������ ������� ������-----------------------------------------------------------
Public Sub GetFireTSquareDataFromAnalize(ByRef shp As Visio.Shape)
'����������� ����� ������� ������� ������ � ����������� � ��������
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---�������� �� ������������ ����� �������� ��� �������
    '���������� ����� ��� ������
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  '���� ��� ����� ������ - ������� �� ��������
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---�������� ������ ������ �� �����������
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---���� � ��������� �������� ��� �� ����� ������ (�.�. ��� �� �����) - ������� �� �����
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "�� ��������� �������� ��� ����� �������� �������! ��������� ������ ������ ����������!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---���������, ������� �� ������ �����
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "������ ����� �����������! ��������� ������ ������ ����������!", vbCritical
        Exit Sub
    End If
    
'---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
'        PS_FireGraph_DeleteKnot shp
        PS_FireTGraph_DeleteKnot shp
    Next i
    
    '---�������� ������ ������
    If shp.Cells("User.IndexPers") = 127 Then
        '���� ������ - ������ ������� ������
        GraphAnalizer.PS_GetFireSquares DataArray
    End If
    
'---�������� ���������� ������ ��������� ���������� ����� �������
    ps_FireTGraphicBuild shp, DataArray
    
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetFireTSquareDataFromAnalize"
End Sub

Public Sub GetFireTSquareDataFromTable(ByRef shp As Visio.Shape)
'����������� ����� ������� ������� ������� � ����������� � �������� ������
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---���� � ����� ����� Cancel, �� ������� �� �����
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_FireGraph_DeleteKnot shp
    Next i
    '---�������� �� ����� ������ ������ ��� ���������� ������� �������� � �������� ������
    DataForm.PS_GetMainArray MainArray
    '---������������� ������
    ps_FireTGraphicBuild shp, MainArray
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetFireTSquareDataFromTable"
End Sub


Private Sub ps_FireTGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray())
'����� ������ ����� ������ ������� ������ (�������)
Dim i As Integer

On Error GoTo EX

    '---������������� �������� ��� ������ (���������) �����
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(MainArray(1, 0)) & "/User.FireMax)*Height"
    
    '---������������� �������� ��� ����� ����� (�������� ��)
    For i = 1 To UBound(MainArray, 2)
        PS_FireTGraph_AddKnot shp
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(MainArray(1, i)) & "/User.FireMax)*Height"
    Next i
    
Exit Sub
EX:
    MsgBox "������ ���� ������� �� ��������� ����� �����������!", vbCritical
    shp.Delete
End Sub



'--------------------------------������ �������-----------------------------------------------------------
Public Sub GetExpenceDataFromAnalize(ByRef shp As Visio.Shape)
'����������� ����� ������� ������� � ����������� � ��������
Dim i As Integer
Dim DataArray()
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---�������� �� ������������ ����� �������� ��� �������
    '���������� ����� ��� ������
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  '���� ��� ����� ������ - ������� �� ��������
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---�������� ������ ������ �� �����������
    '---���������� ����������
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh
    
    '---���� � ��������� �������� ��� �� ����� ������ (�.�. ��� �� �����) - ������� �� �����
    If GraphAnalizer.ColP_Fires.Count = 0 Then
        MsgBox "�� ��������� �������� ��� ����� �������� ������ ����! ��������� ������ ������ ����������!", vbCritical
        Set GraphAnalizer = Nothing
        Exit Sub
    End If
    
    '---���������, ������� �� ������ �����
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "������ ����� �����������! ��������� ������ ������ ����������!", vbCritical
        Exit Sub
    End If
    
'---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_WaterGraph_DeleteKnot shp
    Next i
    
    '---�������� ������ ������
    If shp.Cells("User.IndexPers") = 125 Then
        '���� ������ - ������ �������
        GraphAnalizer.PS_GetWStvolsPodOut DataArray
    ElseIf shp.Cells("User.IndexPers") = 126 Then
        '���� ������ - ������ ������������ �������
        GraphAnalizer.PS_GetWStvolsEffPodOut DataArray
    End If
    
'---�������� ���������� ������ ��������� ���������� ����� �������
    ps_ExpenceGraphicBuild shp, DataArray, True

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetExpenceDataFromAnalize"
End Sub

Public Sub GetExpenceDataFromTable(ByRef shp As Visio.Shape)
'����������� ����� ������� �������� � ����������� � �������� ������
Dim MainArray() As Variant
Dim i As Integer
    
    On Error GoTo EX
    
    DataForm.ShowMe shp
    '---���� � ����� ����� Cancel, �� ������� �� �����
    If DataForm.RefreshNeed = False Then Exit Sub
    
    '---������� ��� ����� ������� ����� ������ (���� �� ����� ������)
    For i = 1 To shp.RowCount(visSectionControls) - 1
        PS_WaterGraph_DeleteKnot shp
    Next i

    '---�������� �� ����� ������ ������ ��� ���������� ������� �������� � �������� ������
    DataForm.PS_GetMainArray MainArray
    '---������������� ������
    ps_ExpenceGraphicBuild shp, MainArray, False
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetExpenceDataFromTable"
End Sub

Private Sub ps_ExpenceGraphicBuild(ByRef shp As Visio.Shape, ByRef MainArray(), ByVal SumOpt As Boolean)
'����� ������ ����� ������ ������� (������������ �������)
'SumOpt: ������ - ���� ����������� ������, ���� - ���� ���������� ��������� ��������
Dim i As Integer
Dim Expence As Double

On Error GoTo EX

    '---������������� �������� ��� ������ (���������) �����
        Expence = MainArray(1, 0)
        shp.Cells("Controls.Row_" & 1).FormulaU = "(" & str(MainArray(0, 0) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & 1 & ".Y").FormulaU = "(" & str(Expence) & "/User.MaxExpense)*Height"
    
    '---������������� �������� ��� ����� ����� (�������� ��)
    For i = 1 To UBound(MainArray, 2)
        PS_WaterGraph_AddKnot shp
        If SumOpt = True Then
            Expence = Expence + MainArray(1, i)  '� ������ ���������� ��� ������� ������
        Else
            Expence = MainArray(1, i)            '� ������ ���������� ��� ������� ������ ����� �������
        End If
        shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(MainArray(0, i) / 60) & "/User.TimeMax)*Width"
        shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(Expence) & "/User.MaxExpense)*Height"
    Next i
    
Exit Sub
EX:
   'Exit
   MsgBox "������� ������ ����������� ������� �� ��������� ����� �����������!", vbCritical
   shp.Delete
End Sub

'--------------------------------���� �������-----------------------------------------------------------
Public Sub GetCommonDataFromAnalize(ByRef shp As Visio.Shape)
'����������� ����� �������� � �������
Dim GraphAnalizer As c_Graph
Dim PageIndex As Integer

    On Error GoTo EX

'---���������� ����������
'---�������� �� ������������ ����� �������� ��� �������
    '���������� ����� ��� ������
    SeetsSelectForm.Show
    If SeetsSelectForm.Flag = False Then Exit Sub  '���� ��� ����� ������ - ������� �� ��������
    PageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).index
    
'---�������� ������ ������ �� �����������
    '---���������� ����������
    Set GraphAnalizer = New c_Graph
    GraphAnalizer.pi_TargetPageIndex = PageIndex
    GraphAnalizer.sC_ColRefresh

'---���������, ������� �� ������ �����
    If GraphAnalizer.PF_CheckFireBeginExist = False Then
        MsgBox "������ ����� �����������! ��������� ������ ������ ����������!", vbCritical
        Exit Sub
    End If

'---�������� �������� � ������
    '!!!�������� ��������� ������ - ����� ����������� �����������!!!
    On Error Resume Next
    '������ ������
'    shp.Cells("Prop.TimeBegin").FormulaU = """" & GraphAnalizer.PF_GetBeginDateTime & """"
    shp.Cells("Prop.TimeBegin").FormulaU = "TheDoc!User.FireTime"
    '������������ ������
'    shp.Cells("Prop.FireMax").FormulaForceU = "Guard(" & GraphAnalizer.PF_GetMaxSquare(GraphAnalizer.ColP_Fires.Count) & ")"
    shp.Cells("Prop.FireMax").FormulaForceU = "Guard(" & GraphAnalizer.GetMaxGraphSize(GraphAnalizer.ColP_Fires.Count) & ")"
    '������������ �����
    shp.Cells("Prop.TimeMax").FormulaForceU = "Guard(" & GraphAnalizer.PF_GetTimeEnd(5, "s") / 60 & ")"
    '����� ���������
    shp.Cells("Prop.TimeEnd").FormulaU = GraphAnalizer.PF_GetTimeEnd(4, "s") / 60
    '�������������
    shp.Cells("Prop.WaterIntense").FormulaForceU = "GUARD(" & str(GraphAnalizer.PF_GetIntence(GraphAnalizer.ColP_Fires.Count)) & ")"
    '����������, ��� ������ ������� ����������� �� ���������� �����
    shp.Cells("Prop.WaterIntenseType").Formula = "INDEX(1;Prop.WaterIntenseType.Format)"
    
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetCommonDataFromAnalize"
End Sub

Public Sub ChangeMaxValues(ByRef shp As Visio.Shape)
'��������� ������������ �������� �������
Dim CurShape As Visio.Shape

    On Error GoTo EX

    '---���������� ��� ������� (�� �����) ������������ �� ���������!!!
    For Each CurShape In Application.ActivePage.Shapes
        If CurShape.CellExists("User.IndexPers", 0) = True And CurShape.CellExists("User.Version", 0) = True Then
            If CurShape.Cells("User.IndexPers") = 123 Or CurShape.Cells("User.IndexPers") = 124 _
                Or CurShape.Cells("User.IndexPers") = 125 Or CurShape.Cells("User.IndexPers") = 126 _
                Or CurShape.Cells("User.IndexPers") = 127 _
                Then   '���� ������ - ������ ����� �������� (� ����������� ��� ����-��)
                
                PS_GraphicsFix CurShape
            End If
        End If
    Next CurShape
    
    '---���������� ����� ��� ������� ������ �������
    MaxValuesForm.PS_ShowME shp

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ChangeMaxValues"
End Sub










