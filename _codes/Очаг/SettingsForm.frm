VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "��������� ����������"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8610
   OleObjectBlob   =   "SettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim matrixSize As Long          '���������� ������ � �������
Dim matrixChecked As Long       '���������� ����������� ������



Private Sub btnStopModelling_Click()
    stopModellingFlag = True
End Sub

Private Sub optTTX_Change()
'    txtNozzleRangeValue.Enabled = False
End Sub



Private Sub optValue_Change()
    txtNozzleRangeValue.Enabled = optValue.value
    If txtNozzleRangeValue.Enabled Then
        txtNozzleRangeValue.BackColor = vbWhite
    Else
        txtNozzleRangeValue.BackColor = &H8000000F
    End If
    
End Sub



'------------------------���������, ���������� �����--------------------------
Private Sub UserForm_Activate()
    Me.txtGrainSize = grain
    matrixSize = 0
    matrixChecked = 0
    
    '���������, �������� �� �������
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "������� ��������. ������ ����� " & grain & "��."
        lblMatrixIsBaked.ForeColor = vbGreen
    Else
        lblMatrixIsBaked.Caption = "������� �� ��������."
        lblMatrixIsBaked.ForeColor = vbRed
    End If
    
    txtGrainSize.value = 200
End Sub



Private Sub btnBakeMatrix_Click()
    '���������� �������� ����� �������
    grain = Me.txtGrainSize
    
    '�������� �������
    MakeMatrix
End Sub

Private Sub btnDeleteMatrix_Click()
    '������� �������
    DestroyMatrix
    
    '���������, ��� ������� �� ��������
    lblMatrixIsBaked.Caption = "������� �� ��������."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRunFireModelling_Click()
'��� ������� �� ������ ��������� �������������
    stopModellingFlag = False
    
    On Error GoTo EX
    '���������� ��������� ���������� �����
    Dim spd As Single
    Dim timeElapsed As Single
    Dim intenseNeed As Single
    spd = Me.txtSpeed
    timeElapsed = Me.txtTime
    intenseNeed = CSng(Replace(Me.txtIntense, ".", ","))
    
    '���������, ��� �� ������ ������� �����
    If timeElapsed > 0 And spd > 0 Then
        '������ �������
        RunFire timeElapsed, spd, intenseNeed
    Else
        MsgBox "�� ��� ������ ��������� �������!", vbCritical
    End If
Exit Sub
EX:
    MsgBox "�� ��� ������ ��������� �������!", vbCritical
End Sub




'--------------------------��������� ���������----------------------------------
Private Function GetMatrixCheckedStatus() As String
'���������� ������� ��� ������� ��������� �������
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    GetMatrixCheckedStatus = "�������� " & procent & "%"
End Function





'--------------------------������� ��������� � �������--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'��������� ��� ����� ����� ���-�� ������ � �������
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long)
'��������� ���-�� ����������� ������
    matrixChecked = matrixChecked + size
    
    '��������� ��������� ������ � ���������� ����������� ������
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub






























