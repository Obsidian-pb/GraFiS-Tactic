VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents PTVAppEvents As Visio.Application
Attribute PTVAppEvents.VB_VarHelpID = -1

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'��������� ��� �������� ���������

'---������� ���������� ����������
Set PTVAppEvents = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'��������� ��� �������� ���������

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---��������� ���������� ���������� ��� ����������� ������������ �� ��������� ����������� �����
    Set PTVAppEvents = Visio.Application
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub PTVAppEvents_CellChanged(ByVal cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer
'---��������� ��� ������
'MsgBox Cell.Name
    
    On Error GoTo Tail
    
    If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
            Or cell.Name = "Prop.StreamType" Or cell.Name = "Prop.Head" Then
        ShpInd = cell.Shape.ID
        '���������, ���������� �� ��������� �������� �� ��
        If cell.Shape.Cells("Prop.TTHType").ResultStr(visString) = "�� ������ ������" Then
            '---��������� ��������� ��������� ������� �������
            If cell.Name = "Prop.TTHType" Then
                StvolModelsListImport (ShpInd)
            End If
            
            '---��������� ��������� ��������� ������� ��������� �������
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Then
                StvolVariantsListImport (ShpInd)
            End If
            
            '---��������� ��������� ��������� ��������� ��� ������ ������� � ������������ � ��� �������
            StvolRFImport (ShpInd)
            
            '---��������� ��������� ��������� ������� ����� ����� �������
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" Then
                StvolStreamTypesListImport (ShpInd)
            End If
            
            '---��������� ��������� ��������� ��������� ������� ������ � ������������ � ��� ����� �����
            StvolDiameterInImport (ShpInd)
            
            '---��������� ��������� ��������� ������ ������� ��� ������� ���� ����� �������
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Then
                StvolHeadListImport (ShpInd)
            End If
            
            '---��������� ��������� ��������� ���� ����� ������ � ������������ � ��� ����� ����� (����������, ����������� ��� ����)
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Then
                StvolStreamValueImport (ShpInd)
            End If
            
            '---��������� ��������� ��������� ������� ���� �� ������
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Or cell.Name = "Prop.Head" Then
                StvolProductionImport (ShpInd)
            End If
            
            
        End If
    End If
    
    '---�������� ������ �� ��������� � wiki-fire.org (������ ��� �������!!!)
    If cell.Name = "Prop.StvolType" Then
        If cell.Shape.Cells("Prop.TTHType").ResultStr(visString) = "�� ������ ������" Then
            '---��������� ����� ��������� ������ �� wiki-fire.org
            StvolWFLinkImport (ShpInd)
        Else
            StvolWFLinkFree (ShpInd)
        End If
    End If
    
    If cell.Name = "Prop.WEType" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� ������� ���������������
        WEModelsListImport (ShpInd)
        '---��������� ��������� ��������� ������� ���� �� ������
        StvolProductionImport (ShpInd)
    End If
    
    If cell.Name = "Prop.WFType" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� ������� ����������� �����
        WFModelsListImport (ShpInd)
        '---��������� ��������� ��������� ������������������ ����������� �����
        StvolProductionImport (ShpInd)
    End If
    
    If cell.Name = "Prop.ColPressure" Or cell.Name = "Prop.Patr" Then
        ShpInd = cell.Shape.ID
        '---��������� ��������� ��������� ������������������ �������
        ColFlowMaxImport (ShpInd)
    End If
    
'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "PTVAppEvents_CellChanged"
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
