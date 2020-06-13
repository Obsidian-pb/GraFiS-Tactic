VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InsertExtSquare 
   Caption         =   "���������� ������� �������"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   OleObjectBlob   =   "F_InsertExtSquare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InsertExtSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fireShape As Visio.Shape            '������ ������� ������



'Public Vfl_TargetShapeID As Long
'Public VmD_TimeStart As Date
'Private vfStr_ObjList() As String


Dim matrixSize As Long              '���������� ������ � �������
Dim matrixChecked As Long           '���������� ����������� ������
'Public timeElapsedMain As Single    '����� ��������� � ������ �������������
'Public pathMain As Single           '���������� ���� � ������ �������������



'--------------------------------���� �������� ��������----------------------------------------------------------------

Private Sub UserForm_Initialize()
'��������� �������� �����

    '---�������� ���������� �������
    FillCBCalculateType

End Sub



Public Function SetFireShape(ByRef shp As Visio.Shape) As F_InsertExtSquare
    Set fireShape = shp
    Set SetFireShape = Me
    
    Me.Caption = "��������� ������� ������� (" & shp.Name & ")"
End Function

Private Sub UserForm_Activate()
'��������� ��������� ����� - ��� ������
    
    '���������, �������� �� �������
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "������� ��������. ������ ����� " & grain & "��."
        lblMatrixIsBaked.ForeColor = vbGreen
        Me.txtGrainSize = grain
    Else
        lblMatrixIsBaked.Caption = "������� �� ��������."
        lblMatrixIsBaked.ForeColor = vbRed
        Me.txtGrainSize.value = 200
    End If
    
End Sub

Private Sub btnRunExtSquareCalc_Click()
'��� ������� �� ������ ��������� ������ � ���������� �������
Dim extSquareCalculator As c_ExtSquareCalculator

    '���������, �������� �� �������
    If Not IsMatrixBacked Then
        MsgBox "������� �� ��������!!!"
        Exit Sub
    End If
    
    '���������� ��������� ����� � ����������
    Set extSquareCalculator = New c_ExtSquareCalculator
    extSquareCalculator.SetOpenSpaceLayer fireModeller
    extSquareCalculator.RunDemon fireShape
    
    Me.Hide
End Sub

Private Sub B_Cancel2_Click()
    Me.Hide
End Sub

Private Sub btnBakeMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "������� ������� ������ �������������� �������! ��������� ������ �������� ����� ��� ����� �������.", vbInformation, "������-������"
        Exit Sub
    End If

    '���������� �������� ����� �������
    grain = Me.txtGrainSize
    
    '�������� �������
    MakeMatrix Me
End Sub

Private Sub btnDeleteMatrix_Click()
    '������� �������
    DestroyMatrix
        
    '���������, ��� ������� �� ��������
    lblMatrixIsBaked.Caption = "������� �� ��������."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRefreshMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "������� ������� ������ �������������� �������! ��������� ������ �������� ����� ��� ����� �������.", vbInformation, "������-������"
        Exit Sub
    End If
    
    '��������� ������� �������� �����������
    RefreshOpenSpacesMatrix Me
End Sub

'--------------------------��������� ��������� �������������----------------------------------
Private Function GetMatrixCheckedStatus() As String
'���������� ������� ��� ������� ��������� �������
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    GetMatrixCheckedStatus = "�������� " & procent & "%"
End Function


'--------------------------������� ��������� � ������� �������������--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'��������� ��� ����� ����� ���-�� ������ � �������
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long)
'��������� ���-�� ����������� ������
    matrixChecked = matrixChecked + size
    
    '��������� ��������� ������ � ����������� ����������� ������
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub


'------------------------���������� �������
Private Sub FillCBCalculateType()
'������ ��������� ��������� �������
    CB_CalculateType.AddItem "��������� ������� �������"
    CB_CalculateType.AddItem "����������� ������� �������"
    CB_CalculateType.AddItem "���� ��������"
    CB_CalculateType.ListIndex = 0
    CB_CalculateType.ControlTipText = "���� �������� - ���� �������� ������ ��������������� ��� ����� ������, �� �������� �� ������������ ����������� �����������;" & Chr(13) & Chr(10) & _
    "��������� ������� ������� - ������ ���������� ������ ��� �������� ������ ������, �������, ���������� � ������������ �������������, �� ���������������;" & Chr(13) & Chr(10) & _
    "����������� ������� ������� - �������������� ������ �� ������� ������ ������, �� ������� ������ ������"
End Sub

Public Property Get AttackDeep() As Byte
    AttackDeep = 5
End Property


