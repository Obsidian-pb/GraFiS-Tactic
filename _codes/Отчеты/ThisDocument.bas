VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vmO_App As Visio.Application
Attribute vmO_App.VB_VarHelpID = -1
Dim ButEvent As c_Buttons
Dim f_CheckForm As MCheckForm





Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    
'---������������ ������ ������������ ������� ������
    Set ButEvent = Nothing
    
'---��������� ���� "������ ��������"
    MCheckForm.CloseThis
    
'---������� ������ � ������ ���������� "�����������"
    DeleteButtons
    
'sP_InfoCollectorDeActivate '������������ ����� InfoCollector ����� ��������� � ������ m_Analize
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
Set vmO_App = Visio.Application
    
'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---���������� ����� InfoCollector ����� ��������� � ������ m_Analize
'    sP_InfoCollectorActivate
    
'---��������� ������ ������������ "�����������"
    AddTB_SpecFunc
    AddButtons
    
'---���������� � ���������� ���� "������ ��������" � ��������� ��������������� ������ (�������� ��� ������ ����� ��-���������)
'    MCheckForm.Show
'    MasterCheckRefresh
    
'---���������� ������ ������������ ������� ������
    Set ButEvent = New c_Buttons
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates
    
End Sub



Private Sub vmO_App_CellChanged(ByVal Cell As IVCell)
'��������� �������� ���� �� �������� ������ Prop.FireMax ��� Prop.TimeMax � ��������� ��������� ���������� �������
Dim vsS_CellName As String

    vsS_CellName = Cell.Name
'    If vsS_CellName = "Prop.FireMax" Or vsS_CellName = "Prop.TimeMax" Then
'        sP_ChangeGraphDirect (Cell.Shape.ID)
'    End If
'    If vsS_CellName = "Prop.FireMax" Or vsS_CellName = "Prop.TimeMax" Or vsS_CellName = "Prop.WaterIntenseH" _
'                                                                Or vsS_CellName = "Prop.WaterIntenseType" Then
'        sP_ChangeGraphDirect (Cell.Shape.ID)
'    End If   ' � �������������� !!!!!!!

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

