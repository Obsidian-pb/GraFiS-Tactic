VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataForm 
   Caption         =   "������� ������"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   OleObjectBlob   =   "DataForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RefreshNeed As Boolean   '����� ������������, ����� �� ��������� ������ ������� ��� ���
Public DataCorrect As Boolean   '����� �� ������� ������
Private TargetShape As Visio.Shape  '������� ������

Const TextFieldWidth As Integer = 42  '������ ���������� ����


Dim Time As String
Dim SqrExp As String
Dim TimeA() As String
Dim SqrExpA() As String

Dim c_TableColumn() As c_TableColumn


Public Sub ShowMe(ByRef shp As Visio.Shape)
'�������� ����� ������ ����� � ������������ �������� �������
Dim i As Integer
Dim IndexPers As Integer

    On Error GoTo EX

    '---�������� ������ ������ (������� � ������ ��������)
    SqrExp = shp.Cells("Scratch.A1").ResultStr(visUnitsString)
    Time = shp.Cells("Scratch.B1").ResultStr(visUnitsString)
    StringToArray SqrExp, ";", SqrExpA()
    StringToArray Time, ";", TimeA()
    
    '---���������� ��� ������ ������� �������������� ���������� � �����������
    If shp.Cells("User.IndexPers") = 123 Or shp.Cells("User.IndexPers") = 124 Then
        '---��� ��������
        Me.Label2 = "������� �.��."
    ElseIf shp.Cells("User.IndexPers") = 125 Or shp.Cells("User.IndexPers") = 126 Then
        '---��� ��������
        Me.Label2 = "������ �/�"
    End If
    '---�������������� ������� ��������� ����������
        ReDim c_TableColumn(UBound(TimeA))
    
    '---��������� �������
        ps_FillTable
    
    
    '---���������� �������������� �����
    Me.Show
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ShowMe"
End Sub




'------------------------------��������� ��������� �������----------------------------------------------------
Private Sub ps_FillTable()
'����� ���������� �������
Dim i As Integer

    On Error GoTo EX

    '---� ����������� � ����������� ����� ������� ����������� ���������� �������� �����
    For i = 0 To UBound(TimeA)
        Set c_TableColumn(i) = New c_TableColumn
        c_TableColumn(i).Activate i, 60 + TextFieldWidth * i, 6, TimeA(i), SqrExpA(i)
        
        '---���������� ������ "��������" � �����
        Me.CB_Add.Left = 60 + TextFieldWidth * (i + 1)

        '---������������� ������ ����� �� ���������� ���������
        If i < 5 Then
            Me.Width = TextFieldWidth * (5) + 36 + 54
        Else
            Me.Width = TextFieldWidth * (i + 1) + 36 + 54
        End If

        '---����������� ������ ���������� �����
        CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
        CB_Cancel.Left = (Me.Width / 2) + 3
    Next i

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ps_FillTable"
End Sub

Private Sub ps_ClearTable()
'����� ������� �������
Dim i As Integer

    On Error GoTo EX

    '---� ����������� � ����������� ����� ������� ����������� ���������� �������� �����
    For i = 0 To UBound(TimeA)
        Set c_TableColumn(i) = Nothing
        
        '---���������� ������ "��������" � ������
        Me.CB_Add.Left = TextFieldWidth * (i + 1) + 6

        '---������������� ������ ����� �� ���������� ���������
        If i < 5 Then
            Me.Width = TextFieldWidth * (5) + 36
        Else
            Me.Width = TextFieldWidth * (i + 1) + 36
        End If

        '---����������� ������ ���������� �����
        CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
        CB_Cancel.Left = (Me.Width / 2) + 3
    Next i

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ps_ClearTable"
End Sub

Public Sub DeleteColumn(ByVal a_index As Integer)
'����� �������� ���������� �������
Dim i As Integer

    On Error GoTo EX

    '---�������� ������ ������ ��� �������/��������
        For i = a_index To UBound(TimeA) - 1
            TimeA(i) = TimeA(i + 1)
            SqrExpA(i) = SqrExpA(i + 1)
        Next i
        ReDim Preserve TimeA(UBound(TimeA) - 1)
        ReDim Preserve SqrExpA(UBound(SqrExpA) - 1)
        
    '---������� �������
        ps_ClearTable
    
    '---�������� ������ �������� ������ ������� �������
        ReDim c_TableColumn(UBound(TimeA))
        
    '---��������� ������� �� �����
        ps_FillTable
        
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "DeleteColumn"
End Sub

Public Sub PS_ChangeTimeValue(ByVal a_ind As Integer, ByVal a_value As String)
    TimeA(a_ind) = a_value
End Sub

Public Sub PS_ChangeDataValue(ByVal a_ind As Integer, ByVal a_value As String)
    SqrExpA(a_ind) = a_value
End Sub

Public Sub PS_GetMainArray(ByRef MainArray())
'������� ����� ��������� ������� ������
Dim i As Integer
Dim NodeCount As Integer
    
    '---���������� ���������� ��������� � ����� �������
        NodeCount = UBound(TimeA)
    
    '---���������� ������ ������ �������
    ReDim MainArray(1, NodeCount)
    '---�������� ������ ������ ��� �������/��������
    For i = 0 To NodeCount
        '---�����
        MainArray(0, i) = TimeA(i) * 60
        '---������
        MainArray(1, i) = SqrExpA(i)
    Next i

End Sub

Private Sub CB_Add_Click()
'��������� ����� �����
Dim lastitem As Integer

    On Error GoTo EX

    lastitem = UBound(c_TableColumn) + 1

    '---����������� ������� ��������
    ReDim Preserve TimeA(lastitem)
    ReDim Preserve SqrExpA(lastitem)
    ReDim Preserve c_TableColumn(lastitem)
    
    TimeA(lastitem) = 0
    SqrExpA(lastitem) = 0

    Set c_TableColumn(lastitem) = New c_TableColumn
    c_TableColumn(lastitem).Activate lastitem, 60 + TextFieldWidth * lastitem, 6, TimeA(lastitem), SqrExpA(lastitem)
    
    '---���������� ������ "��������" � �����
    Me.CB_Add.Left = 60 + TextFieldWidth * (lastitem + 1)

    '---������������� ������ ����� �� ���������� ���������
    If lastitem < 5 Then
        Me.Width = TextFieldWidth * (5) + 36 + 54
    Else
        Me.Width = TextFieldWidth * (lastitem + 1) + 36 + 54
    End If

    '---����������� ������ ���������� �����
    CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
    CB_Cancel.Left = (Me.Width / 2) + 3
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "CB_Add_Click"
End Sub

Public Sub PS_GoToTimeColumn(ByVal index As Integer)
'��������� �������� � ������� ������� � ��������� ��������
'���� ������� � ����� �������� ��� - ����� �� �����
On Error GoTo EX
Dim TB As TextBox

    Set TB = Me.Controls("TB_Time_" & index)

    TB.SetFocus
    TB.SelStart = 0
    TB.SelLength = Len(TB.Text)
    
Exit Sub
EX:
'    Debug.Print "��� ����� �������!"
    MsgBox "����� ������� ���!"
    SaveLog Err, "PS_GoToTimeColumn"
End Sub
Public Sub PS_GoToDataColumn(ByVal index As Integer)
'��������� �������� � ������� ������ � ��������� ��������
'���� ������� � ����� �������� ��� - ����� �� �����
On Error GoTo EX
Dim TB As TextBox

    Set TB = Me.Controls("TB_Data_" & index)
    
    TB.SetFocus
    TB.SelStart = 0
    TB.SelLength = Len(TB.Text)

Exit Sub
EX:
    MsgBox "����� ������� ���!"
    SaveLog Err, "PS_GoToDataColumn"
End Sub

Public Function PF_CheckData() As Boolean
'��������� ���������� ������, ���� � ����� ��� �� ����� ������ � ������� ������
Dim ctrl As Control

    For Each ctrl In Me.Controls
        If ctrl.ForeColor = vbRed Then
            PF_CheckData = False
            Exit Function
        End If
    Next ctrl

PF_CheckData = True
End Function




Private Sub CB_Cancel_Click()
    '---���������, ��� ����� ��������� ������ � �������
    RefreshNeed = False
    '---��������� �����
    CloseForm
End Sub

Private Sub CB_OK_Click()
    '---���������, ��� �� ������ ������� ���������
    If PF_CheckData = False Then
        MsgBox "�� ��� ������ ������� ���������! ���������� �� ��������!", vbCritical
        Exit Sub
    End If
    
    '---���������, ��� ����� ��������� ������ � �������
    RefreshNeed = True
    
    '---��������� �����
    CloseForm
End Sub

Private Sub UserForm_Terminate()
    '---���������, ��� �� ����� ��������� ������ � �������
    RefreshNeed = False
    '---��������� �����
    CloseForm
End Sub

Private Sub CloseForm()
    ps_ClearTable
    Me.Hide
End Sub















Private Sub UserForm_Click()

End Sub




