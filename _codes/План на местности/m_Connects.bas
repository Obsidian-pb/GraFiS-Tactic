Attribute VB_Name = "m_Connects"
Option Explicit
'---------------------------------������ ��� �������� �������� �������� ������-----------------------------

Public Sub Conn(ShpObj As Visio.Shape)
'��������� �������� ������������� �������� � ������� � ������ � ������ ��� ���������
Dim ToShape As Long

'---������������� ��������� ��������� �� ������
On Error Resume Next

'---���� ������� �� � ���� �� ���������, ��������� �������������
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = "Sheet." & ToShape & "!Prop.Street" & ""
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = 0
    End If

'---���������� ������ ������ ������ (�� �������� ����)
    ShpObj.BringToFront
End Sub
