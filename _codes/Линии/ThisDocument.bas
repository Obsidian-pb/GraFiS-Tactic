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

Dim ButEvent As Class1
Dim C_ConnectionsTrace As c_HoseConnector
Dim WithEvents LineAppEvents As Visio.Application
Attribute LineAppEvents.VB_VarHelpID = -1


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'��������� �������� ��������� � �������� ��� ������� ���������

'---������� ������ ButEvent � ������� ������ "�����" � ������ ���������� "�����������"
    Set ButEvent = Nothing
    DeleteButtonLine
    DeleteButtonMLine
    DeleteButtonVHose
    DeleteButtonHoseDrop
    
'---������������ ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then Set app = Nothing
    
'---� ������, ���� �� ������ "����������� ��� �� ����� ������, ������� �
    If Application.CommandBars("�����������").Controls.Count = 0 Then RemoveTBImagination
    
'---������� ���������� ����������
Set LineAppEvents = Nothing
Set C_ConnectionsTrace = Nothing
End Sub


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo EX
'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---��������� ���������� ���������� ��� ����������� ������������ �� ��������� ����������� �����
    Set LineAppEvents = Visio.Application
'---���������� ��������� ������ ��� ������������ ����������
    Set C_ConnectionsTrace = New c_HoseConnector

'---����������� �������
    sp_MastersImport

'---��������� ������ ������ (���� ��� �� ���� ���������)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---������� ������ ���������� "�����������" � ��������� �� ��� ������ "�������� � �����"
    AddTBImagination   '��������� ������� ����������
    AddButtonLine      '��������� ������ ������� �����
    AddButtonMLine     '��������� ������ ������������� �����
    AddButtonVHose     '��������� ������ ����������� �����
    AddButtonHoseDrop  '��������� ������ ��������� � �������

'---���������� ������ ������������ ������� ������
    Set ButEvent = New Class1
    
'---���������� ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

'---���������� ����� ������������ � �������� �����
    Application.ActiveDocument.GlueSettings = visGlueToGeometry + visGlueToGuides + visGlueToConnectionPoints + visGlueToVertices

'---��������� ��� ��������� ������� "FireTime"
    sm_AddFireTime
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates
    
Exit Sub
EX:
    SaveLog Err, "Document_DocumentOpened"
End Sub


Private Sub LineAppEvents_CellChanged(ByVal Cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer
Dim shp As Visio.Shape
Dim Con As Visio.Connect
Dim Shp2 As Visio.Shape
Dim Con2 As Visio.Connect
'---��������� ��� ������
    
    If Cell.Name = "Prop.HoseMaterial" Or Cell.Name = "Prop.HoseDiameter" Then
        ShpInd = Cell.Shape.ID
        '---��������� ��������� ��������� ������� ��������� �������
        HoseDiametersListImport (ShpInd)
        '---��������� ��������� ��������� �������� ������������� �������
        HoseResistanceValueImport (ShpInd)
        '---��������� ��������� ��������� �������� ���������� ����������� �������
        HoseMaxFlowValueImport (ShpInd)
        '---��������� ��������� ��������� �������� ����� �������
        HoseWeightValueImport (ShpInd)
    End If
    
    '���� �������� ������ "User.UseAsRazv" ������ "�����������"
    If Cell.Name = "User.UseAsRazv" Then
        '---��������� ��������� ���������� �������� ���������� ��� ������������
        
        Set shp = Cell.Shape
        
        For Each Con In shp.FromConnects
            Set Shp2 = Con.FromSheet    '���������� ������ - ��� ������� �� ��� ���������� ���������
            For Each Con2 In Shp2.Connects
                C_ConnectionsTrace.Ps_ConnectionAdd Con2
            Next Con2
        Next Con

    End If
    


'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
End Sub

Private Sub sp_TemplatesImport()

Visio.Application.ActivePage.Drop ThisDocument.Masters("����������������"), 0, 0

End Sub


Private Sub sp_MastersImport()
'---����������� �������

    MasterImportSub "����������������"
    MasterImportSub "����������������1000"
    MasterImportSub "�����������������������"
    
End Sub


Private Sub sm_AddFireTime()
'��������� ��������� � �������� �������� - ����� ������ ������, � ������ ��� ����������

    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If

End Sub

Private Sub LineAppEvents_ShapeAdded(ByVal Shape As IVShape)
'������� ���������� �� ���� ������
Dim v_Ctrl As CommandBarControl
Dim SecExists As Boolean
    
'---�������� ��������� ������
    On Error GoTo Tail

'---��������� ����� ������ ������ � � ����������� �� ����� ��������� ��������
    For Each v_Ctrl In Application.CommandBars("�����������").Controls
        If v_Ctrl.State = msoButtonDown Then
            Select Case v_Ctrl.Caption
                Case Is = "�����"
                    '---��������� ��������� ��������� � ������� �������� �����
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---�������� ������ � ������ ������� �������� �����
                            MakeHoseLine Shape, 51, 0
                        End If
                    End If
                Case Is = "����������� �����"
                    '---��������� ��������� ��������� �� ����������� �������� �����
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---�������� ������ � ������ ����������� �������� �����
                            MakeVHoseLine
                        End If
                    End If
                Case Is = "������������� �����"
                    '---��������� ��������� ��������� � ������������� �������� �����
                    If IsSelectedOneShape(False) Then
                    '---���� ������� ���� ���� ������ - �������� �� ��������
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---�������� ������ � ������ ������������� �������� �����
                            MakeHoseLine Shape, 77, 1
                        End If
                    End If
            End Select
        End If
    Next v_Ctrl
    
Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub AddTimeUserCells()
'����� ��������� ������ "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim Cell As Visio.Cell

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

'--------------������������ ��������� ����� ��� ���������� 201� ������
Private Sub app_CellChanged(ByVal Cell As IVCell)
'---���� ��� � ��������� ���������� ������ �� �������
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub

