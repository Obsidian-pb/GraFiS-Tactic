VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ButEventOpenWater As ClassLake
Dim WithEvents WaterAppEvents As Visio.Application
Attribute WaterAppEvents.VB_VarHelpID = -1


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'��������� �������� ��������� � �������� ��� ������� ���������

'---������� ������� ButEvent � WaterAppEvents � ������� ������ "�������� � ������������ ������������" � ������ ���������� "�����������"
    Set ButEventOpenWater = Nothing
    Set WaterAppEvents = Nothing
    DeleteButtons
    
'---� ������, ���� �� ������ "����������� ��� �� ����� ������, ������� �
    If Application.CommandBars("�����������").Controls.Count = 0 Then RemoveTBImagination

End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo Tail
    
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

'---������� ������ ���������� "�����������" � ��������� �� ��� ������ "�������� � ������� ������������"
    AddTBImagination
    AddButtons

'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---���������� ����������� �������
    Set WaterAppEvents = Visio.Application
    Set ButEventOpenWater = New ClassLake
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������."
End Sub


Private Sub WaterAppEvents_CellChanged(ByVal cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer
'---��������� ��� ������
'    ShpInd = Cell.Shape.ID
    
    If cell.Name = "Prop.PipeType" Or cell.Name = "Prop.PipeDiameter" Or cell.Name = "Prop.Pressure" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� ������ ���������
'        DiametersListImport (ShpInd)
        '---��������� ��������� ��������� ������ �������
        PressuresListImport (ShpInd)
        '---��������� ��������� ��������� ����������
        ProductionImport (ShpInd)
    End If

'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
End Sub

Public Sub MastersImport()
'---����������� �������
    MasterImportSub "�������������.vss", "������1_������"
    MasterImportSub "�������������.vss", "������1_�������"
'    MasterImportSub "�������������.vss", "������1_�������" ��_�������
    
End Sub

Private Sub WaterAppEvents_ShapeAdded(ByVal Shape As IVShape)
'������� ���������� �� ���� ������
Dim v_Cntrl As CommandBarControl

'---�������� ��������� ������
    On Error GoTo Tail

'---��������� �������� �� ������ ������� � � ���������� �� ����� �������� ���������� ���������� ������
    Set v_Ctrl = Application.CommandBars("�����������").Controls("������������ ������������")
        If v_Ctrl.State = msoButtonDown Then
'            If v_Ctrl.State = msoButtonUp Then
                If IsSelectedOneShape(False) Then
                '---���� ������� ���� ���� ������ - �������� �� ��������
                    If IsHavingUserSection(False) And IsSquare(False) Then
                    '---�������� ������ � ������ ���� �������
                        ButEventOpenWater.MorphToLake
                    End If
'                Else
'                    PS_CheckButtons v_Ctrl
                End If
'            End If
        
        
        
            '---��������� ��������� ��������� � ������������ ������
'            ButEventOpenWater.MorphToLake
        End If

Set v_Ctrl = Nothing
Exit Sub
Tail:
    SaveLog Err, "WaterAppEvents_ShapeAdded"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������."
    Set v_Ctrl = Nothing
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





