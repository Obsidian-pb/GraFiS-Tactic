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
Public Sub Master_check_refresh()
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

'---��������� ����� ���������
    MCheckForm.ListBox1.Clear

    'GDZS
    If vOC_InfoAnalizer.pi_GDZSpbCount < vOC_InfoAnalizer.pi_GDZSChainsCount Then MCheckForm.ListBox1.AddItem "�� ���������� ����� ������������ ��� ������� ����� ����"
    If vOC_InfoAnalizer.pi_GDZSChainsCount >= 3 And vOC_InfoAnalizer.pi_KPPCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������ ����������-���������� ����� ����"
    'Upravlenie
    If vOC_InfoAnalizer.pi_BUCount >= 3 And vOC_InfoAnalizer.pi_ShtabCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������ ����������� ����"
    If vOC_InfoAnalizer.pi_RNBDCount = 0 And vOC_InfoAnalizer.pi_OchagCount > 0 Then MCheckForm.ListBox1.AddItem "�� ������� �������� �����������"
    If vOC_InfoAnalizer.pi_RNBDCount > 1 Then MCheckForm.ListBox1.AddItem "�������� ���������� ������ ���� �����"
    'Ochag
    If vOC_InfoAnalizer.pi_OchagCount > 0 And vOC_InfoAnalizer.pi_SmokeCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������� ���� ����������"
    If vOC_InfoAnalizer.pi_OchagCount > 0 And vOC_InfoAnalizer.pi_DevelopCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������� ���� ��������������� ������"
    If vOC_InfoAnalizer.pi_SmokeCount > 0 And vOC_InfoAnalizer.pi_OchagCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������ ���� ������"
    If vOC_InfoAnalizer.pi_DevelopCount > 0 And vOC_InfoAnalizer.pi_OchagCount = 0 Then MCheckForm.ListBox1.AddItem "�� ������ ���� ������"
    'PPW
    If vOC_InfoAnalizer.pi_WaterSourceCount > vOC_InfoAnalizer.pi_distanceCount Then MCheckForm.ListBox1.AddItem "�� ������� ���������� �� ������� ������������� �� ����� ������"
    'Hoses
    If vOC_InfoAnalizer.pi_WorklinesCount > vOC_InfoAnalizer.pi_linesPosCount Then MCheckForm.ListBox1.AddItem "�� ������� ��������� (����) ��� ������ ������� �����"
    If vOC_InfoAnalizer.pi_linesCount > vOC_InfoAnalizer.pi_linesLableCount Then MCheckForm.ListBox1.AddItem "�� ������� �������� ��� ������ �������� �����"
    'Plan na mestnosti
    If vOC_InfoAnalizer.pi_BuildCount > vOC_InfoAnalizer.pi_SOCount Then MCheckForm.ListBox1.AddItem "�� ������� ������� ������� ������������� ��� ������� �� ������"
                
                      
    
End Sub
