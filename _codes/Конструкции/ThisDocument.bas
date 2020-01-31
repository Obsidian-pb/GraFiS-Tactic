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
Const cellChangedInterval = 100000

Private ButEvent As c_Buttons


Private Sub app_CellChanged(ByVal Cell As IVCell)
'---���� ��� � ��������� ���������� ������ �� �������
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    OpenDoc
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    CloseDoc
End Sub


Private Sub OpenDoc()
'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---����������� �������
    sp_MastersImport
    
'---������� ������ ���������� "�����������" � ��������� �� ��� ������
    AddTB_Constructions
    AddButtons
    
'---���������� ������ ������������ ������� ������
    Set ButEvent = New c_Buttons
    
'---���������� ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates
End Sub

Private Sub CloseDoc()
'������������ �������� ���������

'---������������ ������ ������������ ������� ������
    Set ButEvent = Nothing
    
'---������������ ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then Set app = Nothing
    
'---������� ������ ���������� "�����������"
    RemoveTB_Constructions
End Sub


Private Sub sp_MastersImport()
'---����������� �������

'---������� 1:200
    MasterImportSub "�����"
    MasterImportSub "�����2"
    MasterImportSub "�����3"
    MasterImportSub "�����4"
    MasterImportSub "���������"
    MasterImportSub "���������2"
    MasterImportSub "�����"
    MasterImportSub "���"
    MasterImportSub "������"
    MasterImportSub "��������������"
'---������� 1:1000
    MasterImportSub "�����_1000"
    MasterImportSub "�����2_1000"
    MasterImportSub "�����3_1000"
    MasterImportSub "�����4_1000"
    MasterImportSub "���������_1000"
    MasterImportSub "���������2_1000"
    
End Sub


