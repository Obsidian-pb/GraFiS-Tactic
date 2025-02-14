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

Dim ButEventOpenWater As ClassLake
Dim WithEvents WaterAppEvents As Visio.Application
Attribute WaterAppEvents.VB_VarHelpID = -1

Private Sub app_CellChanged(ByVal Cell As IVCell)
'---���� ��� � ��������� ���������� ������ �� �������
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEventOpenWater.PictureRefresh
        cellChangedCount = 0
    End If
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'��������� �������� ��������� � �������� ��� ������� ���������

    On Error GoTo Tail

'---������� ������� ButEvent � WaterAppEvents � ������� ������ "�������� � ������������ ������������" � ������ ���������� "�����������"
    Set ButEventOpenWater = Nothing
    Set WaterAppEvents = Nothing
    DeleteButtons
    
'---������������ ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then Set app = Nothing
    
'---� ������, ���� �� ������ "����������� ��� �� ����� ������, ������� �
    If Application.CommandBars("�����������").Controls.Count = 0 Then RemoveTBImagination
    
Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Document_BeforeDocumentClose"

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

'---���������� ����������� �������
    Set WaterAppEvents = Visio.Application
    Set ButEventOpenWater = New ClassLake
    
''---��������� ������� ����������
'    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������.", , ThisDocument.Name
End Sub


Private Sub WaterAppEvents_CellChanged(ByVal Cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer

    On Error GoTo Tail

'---��������� ��� ������
'    ShpInd = Cell.Shape.ID
    
    If Cell.Name = "Prop.PipeType" Or Cell.Name = "Prop.PipeDiameter" Or Cell.Name = "Prop.Pressure" Then
        ShpInd = Cell.Shape.ID
        '---��������� ��������� ��������� ������ ���������
'        DiametersListImport (ShpInd)
        '---��������� ��������� ��������� ������ �������
        PressuresListImport Cell.Shape
        '---��������� ��������� ��������� ����������
        ProductionImport Cell.Shape
    End If

'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������

Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "WaterAppEvents_CellChanged"
End Sub

Public Sub MastersImport()
'---����������� �������
    MasterImportSub "������1_������"
    MasterImportSub "������1_�������"
'    MasterImportSub "������1_�������" ��_�������
    
End Sub

Private Sub WaterAppEvents_ShapeAdded(ByVal Shape As IVShape)
'������� ���������� �� ���� ������
Dim v_Ctrl As CommandBarControl

'---�������� ��������� ������
    On Error GoTo Tail

'---��������� �������� �� ������ ������� � � ���������� �� ����� �������� ���������� ���������� ������
    Set v_Ctrl = Application.CommandBars("�����������").Controls("������������ ������������")
        If v_Ctrl.state = msoButtonDown Then
            If IsSelectedOneShape(False) Then
            '---���� ������� ���� ���� ������ - �������� �� ��������
                If IsHavingUserSection(False) And IsSquare(False) Then
                '---�������� ������ � ������ ���� �������
                    ButEventOpenWater.MorphToLake
                End If
            End If
        End If

Set v_Ctrl = Nothing
Exit Sub
Tail:
    SaveLog Err, "WaterAppEvents_ShapeAdded"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������.", , ThisDocument.Name
    Set v_Ctrl = Nothing
End Sub

Private Sub AddTimeUserCells()
'����� ��������� ������ "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim Cell As Visio.Cell

    On Error GoTo Tail

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If

Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddTimeUserCells"
End Sub





