VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_PressureChange 
   Caption         =   "������ �� �����"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3870
   OleObjectBlob   =   "f_PressureChange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_PressureChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public currentShp As Visio.Shape


Private Sub CB_Cancel_Click()
    Me.Hide
End Sub

Private Sub CB_OK_Click()
    
'---��������� ������ �� �������
    RefreshTBData Me.TB_InPatr1, GetCellName("GFS_InPatrIn1")
    RefreshTBData Me.TB_InPatr2, GetCellName("GFS_InPatrIn2")
    RefreshTBData Me.TB_InWCLeft1, GetCellName("GFS_InWCLeft1")
    RefreshTBData Me.TB_InWCLeft2, GetCellName("GFS_InWCLeft2")
    RefreshTBData Me.TB_InWCRight1, GetCellName("GFS_InWCRight1")
    RefreshTBData Me.TB_InWCRight2, GetCellName("GFS_InWCRight2")
    RefreshTBData Me.TB_Tank1, GetCellName("GFS_InTank1")
    RefreshTBData Me.TB_Tank2, GetCellName("GFS_InTank2")

'---��������� �����
    Me.Hide
End Sub

Private Sub UserForm_Activate()

    CB_OK.SetFocus

    SetTBData Me.TB_InPatr1, GetCellName("GFS_InPatrIn1")
    SetTBData Me.TB_InPatr2, GetCellName("GFS_InPatrIn2")
    SetTBData Me.TB_InWCLeft1, GetCellName("GFS_InWCLeft1")
    SetTBData Me.TB_InWCLeft2, GetCellName("GFS_InWCLeft2")
    SetTBData Me.TB_InWCRight1, GetCellName("GFS_InWCRight1")
    SetTBData Me.TB_InWCRight2, GetCellName("GFS_InWCRight2")
    SetTBData Me.TB_Tank1, GetCellName("GFS_InTank1")
    SetTBData Me.TB_Tank2, GetCellName("GFS_InTank2")
    
End Sub



Private Sub SetTBData(ByRef tb As textBox, ByVal cellName As String)
'����� ������ ��� ���������� ���� �������� ������������ � ��������� ������ �������� ������
'���� ����� ������ ��� - ���� �����������
    On Error GoTo EX
    
    tb.Value = currentShp.Cells(cellName).Result(visNumber)
    tb.Visible = True
    
Exit Sub
EX:
    tb.Value = ""
    tb.Visible = False
End Sub

Private Sub RefreshTBData(ByRef tb As textBox, ByVal cellName As String)
'����� ��������� ������ � ������ Scratch �������� ��������� ��������� �����
    On Error GoTo EX
    
    currentShp.Cells(cellName).FormulaU = Chr(34) & tb.Value & Chr(34)

EX:
End Sub

Private Function GetCellName(ByVal connecionCellName) As String
'������� ���������� ��� ������ �� ���������� ����� ������ � ������ Connections
Dim i As Integer
Dim cll As Visio.cell

    For i = 0 To currentShp.RowCount(visSectionConnectionPts) - 1
        If connecionCellName = currentShp.CellsSRC(visSectionConnectionPts, i, 0).RowNameU Then
            GetCellName = "Scratch.C" & i + 1
            Exit Function
        End If
    Next i
    
EX:
End Function








