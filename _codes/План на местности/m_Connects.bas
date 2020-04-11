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


'��������� ������ ���� ����� ����������
Public Sub DistBuild(ShpObj As Visio.Shape)
Dim FBeg As String, FEnd As String

    FBeg = ShpObj.Cells("BegTrigger").FormulaU
    FEnd = ShpObj.Cells("EndTrigger").FormulaU

'�������
    If InStr(1, FBeg, "��") <> 0 Or InStr(1, FEnd, "��") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "����") <> 0 Or InStr(1, FEnd, "����") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "�������") <> 0 Or InStr(1, FEnd, "�������") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "�����") <> 0 Or InStr(1, FEnd, "�����") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "��") <> 0 Or InStr(1, FEnd, "��") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "������������") <> 0 Or InStr(1, FEnd, "������������") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "�������") <> 0 Or InStr(1, FEnd, "�������") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
'����������
    If InStr(1, FBeg, "������") <> 0 And InStr(1, FEnd, "������") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD(RGB(112, 48, 160))"
  
End Sub
