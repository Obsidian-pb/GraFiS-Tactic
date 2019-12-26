Attribute VB_Name = "m_Analizer"
Option Explicit
'--------------------------------------------------������ ��� ������ � ������� InfoCollector----------------------------
Dim vOC_InfoAnalizer As InfoCollector


Public Sub sP_InfoCollectorActivate()
    Set vOC_InfoAnalizer = New InfoCollector
End Sub

Public Sub sP_InfoCollectorDeActivate()
    Set vOC_InfoAnalizer = Nothing
End Sub

Public Sub sP_ChangeValue(ShpObj As Visio.Shape)
'��������� ������� �� �������� ������������
Dim i As Integer
Dim psi_TargetPageIndex As Integer

'---���������� ������������ ������� �������� ��� �������
    SeetsSelectForm.Show
    psi_TargetPageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).Index

'---�������� ��������� �������� �������
'    psi_TargetPageIndex = ActivePage.Index
    If vOC_InfoAnalizer Is Nothing Then
        Set vOC_InfoAnalizer = New InfoCollector
    End If
    vOC_InfoAnalizer.sC_Refresh (psi_TargetPageIndex)

'---��������� ����� ��������� ����� ������
    sP_ChangeValueMain ShpObj.ID, psi_TargetPageIndex

End Sub

Public Sub sP_ChangeValueMain(asi_ShpInd As Integer, asi_TargetPage As Integer)
'��������� ������� �� �������� ������������
Dim i As Integer
Dim vsO_TargetPage As Visio.Page
Dim vsO_BaseShape As Visio.Shape
Dim vsO_Shape As Visio.Shape

    On Error GoTo EX
'---���������� ������� ������
    Set vsO_TargetPage = Application.ActiveDocument.Pages(asi_TargetPage)
    Set vsO_BaseShape = Application.ActivePage.Shapes.ItemFromID(asi_ShpInd)

'---���������� ��� ������ � ������
    For i = 1 To vsO_BaseShape.Shapes.Count
        Set vsO_Shape = vsO_BaseShape.Shapes(i)
        If vsO_Shape.CellExists("Actions.ChangeValue", 0) = True Then '---��������� �������� �� ������ �������!!!
            sP_ChangeValueMain vsO_Shape.ID, asi_TargetPage
        End If
        If vsO_Shape.CellExists("User.PropertyValue", 0) = True Then '---��������� �������� �� ������ ����� ������
            vsO_Shape.Cells("User.PropertyValue").FormulaU = _
                str(fp_SetValue(vsO_Shape.Cells("Prop.PropertyName").ResultStr(visUnitsString)))
        End If
    Next i

Set vsO_Shape = Nothing
Set vsO_BaseShape = Nothing
Exit Sub
EX:
    Set vsO_Shape = Nothing
    Set vsO_BaseShape = Nothing
    SaveLog Err, "sP_ChangeValueMain"
End Sub



Private Function fp_SetValue(ass_PropertyName As String) As Double
'��������� ������������� � ���� �������� ������ �������� ��������
'---��������� �������� ������ ������ �������� �������� ��������
'fp_SetValue = 111
Select Case ass_PropertyName
    
    '---�����-------------
    Case Is = "�������� ��"
        fp_SetValue = vOC_InfoAnalizer.pi_MainPAHave + vOC_InfoAnalizer.pi_TargetedPAHave
    Case Is = "�������� ��"
        fp_SetValue = vOC_InfoAnalizer.pi_ALCount + vOC_InfoAnalizer.pi_AKPCount
    Case Is = "������� ��"
        fp_SetValue = vOC_InfoAnalizer.pi_ACCount
    Case Is = "������� �����"
        fp_SetValue = vOC_InfoAnalizer.pi_AGCount
    Case Is = "������� �����������"
        fp_SetValue = vOC_InfoAnalizer.pi_ALCount
    Case Is = "������� ���������������"
        fp_SetValue = vOC_InfoAnalizer.pi_AKPCount
    Case Is = "������� ������� ���"
        fp_SetValue = vOC_InfoAnalizer.pi_MVDCount
    Case Is = "������� ������� ��������"
        fp_SetValue = vOC_InfoAnalizer.pi_MZdravCount
    Case Is = "������ ��������"
        fp_SetValue = vOC_InfoAnalizer.pi_BUCount
    Case Is = "������� ����"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount
    Case Is = "������� �������� ������"
        fp_SetValue = vOC_InfoAnalizer.pi_FireTotalCount
    Case Is = "������� �� ���"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount
    Case Is = "������� �� ��� (������)"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount
    
    '---�����-------------
    
    
    Case Is = "�������� �� ������ ����������"
        fp_SetValue = vOC_InfoAnalizer.pi_MainPAHave
    Case Is = "��������� ��"
        fp_SetValue = PF_RoundUp((vOC_InfoAnalizer.pi_PersonnelNeed + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3) * 3) / 4)
    Case Is = "��������� ���"
        fp_SetValue = PF_RoundUp((vOC_InfoAnalizer.pi_PersonnelNeed + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3) * 3) / 5)
    Case Is = "�������� �� �������� ����������"
        fp_SetValue = vOC_InfoAnalizer.pi_TargetedPAHave
    Case Is = "����������� ��"
        fp_SetValue = vOC_InfoAnalizer.pi_SpecialPAHave
    Case Is = "������ �������"
        fp_SetValue = vOC_InfoAnalizer.pi_OtherTechincsHave
    Case Is = "������� ������� �������"
        fp_SetValue = vOC_InfoAnalizer.pi_PersonnelHave
        
    Case Is = "��������� ������� �������" '� ������ ��������� �������
        fp_SetValue = vOC_InfoAnalizer.pi_PersonnelNeed + vOC_InfoAnalizer.pi_GDZSMansRezCount
'        + Int(vOC_InfoAnalizer.pi_HosesCount * 20 / 100)  ' �������� �����!!!!!
            
    Case Is = "����������� ���������� ������� ����"
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount
    Case Is = "��������� ��������� �������"
'        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3)
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_GDZSChainsRezCount)
    Case Is = "��������� ������� ����"
'        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3)
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.ps_GDZSChainsRezCount)
    Case Is = "����������� ���������� ������������������"
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSMansCount
    Case Is = "������� ������"
        fp_SetValue = vOC_InfoAnalizer.ps_FireSquare
    Case Is = "������� �������"
        fp_SetValue = vOC_InfoAnalizer.ps_ExtSquare
    Case Is = "��������� ������ ����"
        fp_SetValue = vOC_InfoAnalizer.ps_NeedStreemW
    Case Is = "����������� ������ ����"
        fp_SetValue = vOC_InfoAnalizer.ps_FactStreemW
    Case Is = "������ ������� �������"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWHave
    Case Is = "������ ������ �������"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolFoamHave
    Case Is = "������ ���������� �������"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolPowderHave
    Case Is = "������ ������� �������"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolGasHave
    Case Is = "������ ������� �"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWAHave
    Case Is = "������ ������� �"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWBHave
    Case Is = "������ �������� �������"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWLHave


    Case Is = "��������� ������ ������� �"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 3.7)
    Case Is = "��������� ������ ������� �"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 7.4)
    Case Is = "��������� ������ �������� �������"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 12)
        
    Case Is = "���������� ����"
        fp_SetValue = vOC_InfoAnalizer.ps_GetedWaterValue
    Case Is = "�������� ���������� ����"
        fp_SetValue = vOC_InfoAnalizer.ps_GetedWaterValueMax
    Case Is = "����������� �� �������������"
        fp_SetValue = vOC_InfoAnalizer.pi_GetingWaterCount
    Case Is = "��������� ���������� �� ������������� ��"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32)

    Case Is = "���������� ������� 51��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses51Count
    Case Is = "���������� ������� 66��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses66Count
    Case Is = "���������� ������� 77��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses77Count
    Case Is = "���������� ������� 89��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses89Count
    Case Is = "���������� ������� 110��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses110Count
    Case Is = "���������� ������� 150��"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses150Count
    Case Is = "����� ����� �������� �����"
        fp_SetValue = vOC_InfoAnalizer.pi_HosesLength
        
    Case Is = "����������� ����� ����"
        fp_SetValue = vOC_InfoAnalizer.pi_WaterValueHave
    Case Is = "��������� ����� ���� (10���)"
        fp_SetValue = vOC_InfoAnalizer.ps_FactStreemW * 600   '��� 10 �����

End Select


End Function

'=================================== MASTER CHECK by Vasilchenko ================================================
Public Sub MasterCheckRefresh()
'��������� ������� �� �������� ������������
Dim i As Integer
Dim psi_TargetPageIndex As Integer

    psi_TargetPageIndex = Application.ActivePage.Index

'---�������� ��������� �������� �������
'    psi_TargetPageIndex = ActivePage.Index
    If vOC_InfoAnalizer Is Nothing Then
        Set vOC_InfoAnalizer = New InfoCollector
    End If
    vOC_InfoAnalizer.sC_Refresh (psi_TargetPageIndex)

'---��������� ������� ���������
    MCheckForm.ListBox1.Clear
    MCheckForm.ListBox2.Clear
    Dim comment As Boolean
    comment = False
    'Ochag
    If vOC_InfoAnalizer.pi_OchagCount = 0 Then
        If vOC_InfoAnalizer.pi_SmokeCount > 0 Or vOC_InfoAnalizer.pi_DevelopCount > 0 Or vOC_InfoAnalizer.pi_FireCount Then
            MCheckForm.ListBox1.AddItem "�� ������ ���� ������"
            comment = True
        End If
    End If
    If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_SmokeCount = 0 Then
        MCheckForm.ListBox1.AddItem "�� ������� ���� ����������"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_DevelopCount = 0 Then
        MCheckForm.ListBox1.AddItem "�� ������� ���� ��������������� ������"
        comment = True
    End If
    'Upravlenie
    If vOC_InfoAnalizer.pi_BUCount >= 3 And vOC_InfoAnalizer.pi_ShtabCount = 0 Then
        MCheckForm.ListBox1.AddItem "�� ������ ����������� ����"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_RNBDCount = 0 And vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 Then
        MCheckForm.ListBox1.AddItem "�� ������� �������� �����������"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_RNBDCount > 1 Then
        MCheckForm.ListBox1.AddItem "�������� ���������� ������ ���� �����"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_BUCount >= 5 And vOC_InfoAnalizer.pi_SPRCount <= 1 Then
        MCheckForm.ListBox1.AddItem "�� ������������ ������� ���������� �����"
        comment = True
    End If
    'GDZS
    If vOC_InfoAnalizer.pi_GDZSpbCount < vOC_InfoAnalizer.pi_GDZSChainsCount Then
        MCheckForm.ListBox1.AddItem "�� ���������� ����� ������������ ��� ������� ����� ���� (" & vOC_InfoAnalizer.pi_GDZSpbCount & "/" & vOC_InfoAnalizer.pi_GDZSChainsCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsCount >= 3 And vOC_InfoAnalizer.pi_KPPCount = 0 Then
        MCheckForm.ListBox1.AddItem "�� ������ ����������-���������� ����� ����"
        comment = True
    End If
    If vOC_InfoAnalizer.pb_GDZSDiscr = True Then
        MCheckForm.ListBox1.AddItem "� ������� �������� ������ ���� ������ �������� �� ����� ��� �� ���� ������������������"
        comment = True
    End If
    If Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) - vOC_InfoAnalizer.pi_GDZSChainsRezCount <> 0 Then
        MCheckForm.ListBox1.AddItem "������������ ��������� ������� ���� (" & vOC_InfoAnalizer.pi_GDZSChainsRezCount & "/" & Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) & ")"
        comment = True
    End If
    'PPW
    If vOC_InfoAnalizer.pi_WaterSourceCount > vOC_InfoAnalizer.pi_distanceCount Then
        MCheckForm.ListBox1.AddItem "�� ������� ���������� �� ������� ������������� �� ����� ������ (" & vOC_InfoAnalizer.pi_distanceCount & "/" & vOC_InfoAnalizer.pi_WaterSourceCount & ")"
        comment = True
    End If
    'Hoses
'    If vOC_InfoAnalizer.pb_AllHosesWithPos Then MCheckForm.ListBox1.AddItem "�� ������� ��������� (����) ��� ������ ������� �����"
    If vOC_InfoAnalizer.pi_WorklinesCount > vOC_InfoAnalizer.pi_linesPosCount Then
        MCheckForm.ListBox1.AddItem "�� ������� ��������� (����) ��� ������ ������� ����� (" & vOC_InfoAnalizer.pi_linesPosCount & "/" & vOC_InfoAnalizer.pi_WorklinesCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_linesCount > vOC_InfoAnalizer.pi_linesLableCount Then
        MCheckForm.ListBox1.AddItem "�� ������� �������� ��� ������ �������� ����� (" & vOC_InfoAnalizer.pi_linesLableCount & "/" & vOC_InfoAnalizer.pi_linesCount & ")"
        comment = True
    End If
    'Plan na mestnosti
    If vOC_InfoAnalizer.pi_BuildCount > vOC_InfoAnalizer.pi_SOCount Then
        MCheckForm.ListBox1.AddItem "�� ������� ������� ������� ������������� ��� ������� �� ������ (" & vOC_InfoAnalizer.pi_SOCount & "/" & vOC_InfoAnalizer.pi_BuildCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_OrientCount = 0 And vOC_InfoAnalizer.pi_BuildCount > 0 Then
        MCheckForm.ListBox1.AddItem "�� ������� ��������� �� ���������, ����� ��� ���� ������ ��� ������� �����"
        comment = True
    End If
            '����� ��������� ������
         
    If vOC_InfoAnalizer.ps_FactStreemW <> 0 And vOC_InfoAnalizer.ps_FactStreemW < vOC_InfoAnalizer.ps_NeedStreemW Then
        MCheckForm.ListBox1.AddItem "������������� ����������� ������ ���� (" & vOC_InfoAnalizer.ps_FactStreemW & " �/c < " & vOC_InfoAnalizer.ps_NeedStreemW & " �/�)"
        comment = True
    End If
    If (vOC_InfoAnalizer.ps_FactStreemW * 600) > vOC_InfoAnalizer.pi_WaterValueHave Then
        If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) > vOC_InfoAnalizer.pi_GetingWaterCount Then MCheckForm.ListBox1.AddItem "������������� ������������� ������ �������" '& (" & vOC_InfoAnalizer.pi_GetingWaterCount & "/" & PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_PersonnelHave < vOC_InfoAnalizer.pi_PersonnelNeed Then
        MCheckForm.ListBox1.AddItem "������������ ������� �������, � ������ ��������� ������� (" & vOC_InfoAnalizer.pi_PersonnelHave & "/" & vOC_InfoAnalizer.pi_PersonnelNeed & ")"
        comment = True
    End If
    If comment = False Then MCheckForm.ListBox1.AddItem "��������� �� ����������"

    '============������ ������� - ������ ����������� ������===========
    If vOC_InfoAnalizer.pi_TechTotalCount <> 0 Then
        MCheckForm.ListBox2.AddItem "������� ����"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_TechTotalCount
    End If
    If vOC_InfoAnalizer.pi_MVDCount <> 0 Then
        MCheckForm.ListBox2.AddItem "������� ���"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MVDCount
    End If
    If vOC_InfoAnalizer.pi_MZdravCount <> 0 Then
        MCheckForm.ListBox2.AddItem "������� ��������"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MZdravCount
    End If
    If vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount <> 0 Then
        MCheckForm.ListBox2.AddItem "������� ���� ��������"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount
    End If
    If vOC_InfoAnalizer.pi_FireTotalCount <> 0 Then
         MCheckForm.ListBox2.AddItem "������� �������� ������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_FireTotalCount
    End If
    If vOC_InfoAnalizer.pi_TechTotalCount <> 0 Then
         MCheckForm.ListBox2.AddItem "�������� ��"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MainPAHave + vOC_InfoAnalizer.pi_TargetedPAHave & " (" & vOC_InfoAnalizer.pi_TargetedPAHave & " ���.����., " & vOC_InfoAnalizer.pi_MainPAHave & " ���.����.)"
    End If
    If vOC_InfoAnalizer.pi_GetingWaterCount <> 0 Then
         MCheckForm.ListBox2.AddItem "������������� ��������������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GetingWaterCount
    End If
    If vOC_InfoAnalizer.pi_SpecialPAHave <> 0 Then
         MCheckForm.ListBox2.AddItem "����������� ��"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_SpecialPAHave & " (" & vOC_InfoAnalizer.pi_ALCount + vOC_InfoAnalizer.pi_AKPCount & " ��������)"
    End If
    If vOC_InfoAnalizer.pi_OtherTechincsHave <> 0 Then
         MCheckForm.ListBox2.AddItem "������ ���.�������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_OtherTechincsHave
    End If
    If vOC_InfoAnalizer.pi_BUCount <> 0 Then
         MCheckForm.ListBox2.AddItem "������ ��������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_BUCount
    End If
    If vOC_InfoAnalizer.pi_SPRCount <> 0 Then
         MCheckForm.ListBox2.AddItem "�������� ���������� �����"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_SPRCount
    End If
    If vOC_InfoAnalizer.pi_PersonnelHave <> 0 Then
         MCheckForm.ListBox2.AddItem "������� ������� (��� ���������)"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_PersonnelHave
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsCount <> 0 Then
         MCheckForm.ListBox2.AddItem "�������� ������� ����"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GDZSChainsCount & " (" & vOC_InfoAnalizer.pi_GDZSMansCount & " ������������������)"
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsRezCount <> 0 Then
         MCheckForm.ListBox2.AddItem "������� ���� � �������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GDZSChainsRezCount & " (" & vOC_InfoAnalizer.pi_GDZSMansRezCount & " ������������������)"
    End If
    If vOC_InfoAnalizer.ps_FireSquare <> 0 Then
         MCheckForm.ListBox2.AddItem "������� ������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_FireSquare & " � ��. (������� ������� " & vOC_InfoAnalizer.ps_ExtSquare & " � ��.)"
    End If
    If vOC_InfoAnalizer.pi_StvolWHave <> 0 Then
         MCheckForm.ListBox2.AddItem "������� �������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolWHave & " (" & vOC_InfoAnalizer.pi_StvolWBHave & " ���.�, " & vOC_InfoAnalizer.pi_StvolWAHave & " ���.�, " & vOC_InfoAnalizer.pi_StvolWLHave & " ��������)"
    End If
    If vOC_InfoAnalizer.pi_StvolFoamHave <> 0 Then
         MCheckForm.ListBox2.AddItem "������ �������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolFoamHave
    End If
    If vOC_InfoAnalizer.pi_StvolPowderHave <> 0 Then
         MCheckForm.ListBox2.AddItem "���������� �������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolPowderHave
    End If
    If vOC_InfoAnalizer.pi_StvolGasHave <> 0 Then
         MCheckForm.ListBox2.AddItem "������ ������� �������"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolGasHave
    End If
    If vOC_InfoAnalizer.ps_FactStreemW <> 0 Then
         MCheckForm.ListBox2.AddItem "����������� ������ ����"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_FactStreemW & " �/�"
           
    End If
    If vOC_InfoAnalizer.ps_NeedStreemW <> 0 Then
         MCheckForm.ListBox2.AddItem "��������� ������ ����"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_NeedStreemW & " �/�"
    End If
    If vOC_InfoAnalizer.pi_WaterValueHave <> 0 Then
         MCheckForm.ListBox2.AddItem "����� ���� � �������� ��"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_WaterValueHave / 1000 & " �"
    End If
    If vOC_InfoAnalizer.pi_linesCount - vOC_InfoAnalizer.pi_WorklinesCount <> 0 Then
        MCheckForm.ListBox2.AddItem "������������� �����"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_linesCount - vOC_InfoAnalizer.pi_WorklinesCount
    End If
    If vOC_InfoAnalizer.pi_HosesLength <> 0 Then
        MCheckForm.ListBox2.AddItem "����� ����� �������� �����"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_HosesLength & " �"
    End If
    If vOC_InfoAnalizer.pi_Hoses51Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 51 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses51Count
    End If
    If vOC_InfoAnalizer.pi_Hoses66Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 66 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses66Count
    End If
    If vOC_InfoAnalizer.pi_Hoses77Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 77 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses77Count
    End If
    If vOC_InfoAnalizer.pi_Hoses89Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 89 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses89Count
    End If
    If vOC_InfoAnalizer.pi_Hoses110Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 110 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses110Count
    End If
    If vOC_InfoAnalizer.pi_Hoses150Count <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ������� 150 ��"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses150Count
    End If
    If vOC_InfoAnalizer.ps_GetedWaterValue <> 0 Then
        MCheckForm.ListBox2.AddItem "���������� ����"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_GetedWaterValue & " �/� (" & "max = " & vOC_InfoAnalizer.ps_GetedWaterValueMax & " �/�)"
    End If
    If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) <> 0 Then
        MCheckForm.ListBox2.AddItem "��������� ���������� ��-40 �� ����� (�� �������)"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32)
    End If
    If vOC_InfoAnalizer.ps_FactStreemW * 600 <> 0 Then
        MCheckForm.ListBox2.AddItem "��������� ����� ���� (�� �������, �� 10 ���)"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = (vOC_InfoAnalizer.ps_FactStreemW * 600) / 1000 & " �"
    End If
        
End Sub
