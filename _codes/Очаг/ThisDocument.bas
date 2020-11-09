VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 1000

Dim WithEvents SquareAppEvents As Visio.Application
Attribute SquareAppEvents.VB_VarHelpID = -1

Dim ButEventFireArea As ClassFireArea, ButEventStorm As ClassStorm, ButEventFog As ClassFog, ButEventRush As ClassRush

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo EX
    
'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---����������� �������
    MastersImport

'---��������� ������ ������ (���� ��� �� ���� ���������)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---���������� ������ SquareAppEvents ��� ������������ �� �������� ������������
    Set SquareAppEvents = Visio.Application
    
'---������� ������ ���������� "�����������" � ��������� �� ��� ������ "�������� � ���� �������"
    AddTBImagination
    AddButtons
    
'---���������� ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---�������� ����� ����
    VfB_NotShowPropertiesWindow = False

'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

Set ButEventFireArea = New ClassFireArea
Set ButEventStorm = New ClassStorm
Set ButEventFog = New ClassFog
Set ButEventRush = New ClassRush

'---��������� �������� ��������� "FireTime"
    sm_AddFireTime
Exit Sub
EX:
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'��������� �������� ��������� � �������� ��� ������� ���������

'---������� ������ ButEvent � ������� ������ "�������" � ������ ���������� "�����������"
    Set ButEventFireArea = Nothing
    Set ButEventStorm = Nothing
    Set ButEventFog = Nothing
    Set ButEventRush = Nothing
    DeleteButtons
    
'---������������ ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then Set app = Nothing
    
'---� ������, ���� �� ������ "����������� ��� �� ����� ������, ������� �
    If Application.CommandBars("�����������").Controls.Count = 0 Then RemoveTBImagination
'---������� ���������� ����������
    Set SquareAppEvents = Nothing
    
End Sub


Private Sub SquareAppEvents_CellChanged(ByVal cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Long '(64) - ������� ������
'Dim shpFire As Visio.Shape '������ ������� ������

'---��������� �� ��������� �� ������� � �������
    On Error GoTo EX
    If cell.Shape.ContainingMasterID >= 0 Then Exit Sub
    
'---��������� ��� ������
    If cell.Name = "Prop.FireCategorie" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� ������� �������� �������� ������ ��� ��������� ���������
        DescriptionsListImport (ShpInd)
    End If
        
    If cell.Name = "Prop.FireDescription" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� �������� �������� ������ ��� ������� ��������
        GetFactorsByDescription (ShpInd)
    End If
    
    If cell.Name = "Prop.FireTime" Then
        '---��������� ����� ������ �� ���� ������ ������ � ���� ���� ���������
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
            "DATETIME(" & str(CDbl(cell.Shape.Cells("Prop.FireTime").Result(visDate))) & ")"
    End If
        
        
'    ElseIf Cell.Name = "Prop.FireObject" Then
'        '---��������� ��������� ��������� �������� �������������� ������ ���� �������� ������
'        GetIntenseWaterByObject (ShpInd)
'        '---��������� ��������� ��������� �������� �������� �������� ��� �������� ������
'        GetSpeedByObject (ShpInd)
'    ElseIf Cell.Name = "Prop.FireMaterials" Then
'        '---��������� ��������� ��������� �������� �������������� ������ ���� ������� ����������
'        GetIntenseWaterByMaterial (ShpInd)
'        '---��������� ��������� ��������� �������� �������� �������� ��� ���������� ������
'        GetSpeedByMaterial (ShpInd)
'    End If
    
    
    
'MsgBox Cell.Shape.Index
'MsgBox Cell.Shape.ID

'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
EX:
End Sub

Public Sub MastersImport()
'---����������� �������
'Dim mstr As Visio.Master

    MasterImportSub "����������1_������"
    MasterImportSub "����������2_������"
    MasterImportSub "����������3_������"
    MasterImportSub "����������4_������"
    MasterImportSub "����������5_������"
    MasterImportSub "����������6_������"
    MasterImportSub "����������1_�������"
    MasterImportSub "����������2_�������"
    MasterImportSub "����������3_�������"
    MasterImportSub "����������4_�������"
    MasterImportSub "����������5_�������"
    MasterImportSub "����������6_�������"
    MasterImportSub "����������1_�������"
    MasterImportSub "����������2_�������"
    MasterImportSub "����������3_�������"
    MasterImportSub "����������4_�������"
    MasterImportSub "����������5_�������"
    MasterImportSub "����������6_�������"
    MasterImportSub "����1_������"
    MasterImportSub "����2_������"
    MasterImportSub "����3_������"
    MasterImportSub "����4_������"
    MasterImportSub "����1_�������"
    MasterImportSub "����2_�������"
    MasterImportSub "����3_�������"
    MasterImportSub "����4_�������"
    MasterImportSub "����1_�������"
    MasterImportSub "����2_�������"
    MasterImportSub "����3_�������"
    MasterImportSub "����4_�������"
    MasterImportSub "�������� �����"
    MasterImportSub "���������"

End Sub



Private Sub sm_AddFireTime()
'��������� ��������� � �������� �������� - ����� ������ ������, � ������ ��� ����������

    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If

End Sub

Private Sub SquareAppEvents_ShapeAdded(ByVal Shape As IVShape)
'������� ���������� �� ���� ������
Dim v_Ctrl As CommandBarControl
'Dim SecExists As Boolean
    
'---�������� ��������� ������
    On Error GoTo Tail

'---��������� ����� ������ ������ � � ����������� �� ����� ��������� ��������
    For Each v_Ctrl In Application.CommandBars("�����������").Controls
        If v_Ctrl.State = msoButtonDown Then
            Select Case v_Ctrl.Caption
                Case Is = "�������"
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---��������� ��������� ��������� � ���� �������
                        ButEventFireArea.MorphToFireArea
                        End If
                    End If
                Case Is = "����������"
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---��������� ��������� ��������� � ����������
                        ButEventFog.MorphToFog
                        End If
                    End If
                Case Is = "���������"
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---��������� ��������� ��������� � ���������
                        ButEventRush.MorphToRush
                        End If
                    End If
                Case Is = "�����"
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---��������� ��������� ��������� � �����
                        ButEventStorm.MorphToStorm
                        End If
                    End If
            End Select
        End If
    Next v_Ctrl
    
Exit Sub
Tail:
    MsgBox "� ���� ������ ��������� �������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub AddTimeUserCells()
'����� ��������� ������ "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If

End Sub

Private Sub app_CellChanged(ByVal cell As IVCell)
'---���� ��� � ��������� ���������� ������ �� �������
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEventFireArea.PictureRefresh
        ButEventStorm.PictureRefresh
        ButEventFog.PictureRefresh
        ButEventRush.PictureRefresh
        cellChangedCount = 0
    End If
End Sub


