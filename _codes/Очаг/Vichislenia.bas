Attribute VB_Name = "Vichislenia"
Option Explicit

Sub SquareSet(ShpObj As Visio.Shape)
'��������� ���������� ���������� ���� ���������� ������ �������� ������� ������
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 '��������� �� ���������� ������ � ���������� �����
ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub


Sub s_SetFireTime(ShpObj As Visio.Shape, Optional showDoCmd As Boolean = True)
'��������� ���������� ������ ��������� User.FireTime �������� ������� ���������� ��� ����������� ������ "����"
Dim vD_CurDateTime As Double

On Error Resume Next

'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---����������� �������� ������� ������������� ������ ������� ��������
            vD_CurDateTime = Now()
            ShpObj.Cells("Prop.FireTime").FormulaU = _
                "DATETIME(" & str(vD_CurDateTime) & ")"
        
        '---���������� ���� ������� ������
            If showDoCmd Then Application.DoCmd (1312)
            
        '---���� � ����-����� ��������� ����������� ������ "User.FireTime", ������� �
            If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
                Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
            End If
            
        '---��������� ����� ������ �� ���� ������ ������ � ���� ���� ���������
            Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
                "DATETIME(" & str(CDbl(ShpObj.Cells("Prop.FireTime").Result(visDate))) & ")"
    Else
        '---���������� ���� ������� ������
            If showDoCmd Then Application.DoCmd (1312)
    End If

End Sub
