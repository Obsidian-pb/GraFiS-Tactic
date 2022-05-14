Attribute VB_Name = "m_FireShapeInsert"
Option Explicit
'--------------------------------------������ ���������� �������� ������� �� ��������� ������----------------------


Public Sub Sm_ShapeFormShow(ShpObj As Visio.Shape)
'��������� ��������� ����� ���������� �������� ������� � ������������ � ��������� ������������
Dim timeStart As Date
Dim time1Stvol As Date

    On Error GoTo EX
'---���������� ��������� �������� �����
    timeStart = ActiveDocument.DocumentSheet.Cells("User.FireTime").ResultStr(0)
    time1Stvol = CellVal(Application.ActiveDocument.DocumentSheet, "User.FirstStvolTime", visDate)
'    F_InsertFire.TB_Time.Value = ShpObj.Cells("Prop.FireTime").ResultStr(visDate)
    F_InsertFire.TB_Time.value = ActiveDocument.DocumentSheet.Cells("User.CurrentTime").ResultStr(0)
    F_InsertFire.TB_Duration.value = DateDiff("n", timeStart, _
                                        ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate))
    F_InsertFire.TB_Radius.value = Round(ShpObj.Shapes.item(4).Cells("Width").Result(visMeters), 2)

'---��������� ������� �����, � ����� ID ������  ��� ������
    F_InsertFire.Vfl_TargetShapeID = ShpObj.ID

'---��������� ������� �����, ��������� ����
'    F_InsertFire.VmD_TimeStart = ActiveDocument.DocumentSheet.Cells("User.FireTime").Result(visDate)
    F_InsertFire.VmD_TimeStart = timeStart
    F_InsertFire.FireTime.Caption = "������ ������: " & ActiveDocument.DocumentSheet.Cells("User.FireTime").ResultStr(0)
'---��������� ������� �����, ���� ������ 1 ������
    If Not time1Stvol = 0 Then
        F_InsertFire.VmD_Time1Stvol = time1Stvol
        F_InsertFire.FireTime.Caption = F_InsertFire.FireTime.Caption & _
            " | ������ 1 ������: " & CellVal(Application.ActiveDocument.DocumentSheet, "User.FirstStvolTime", visUnitsString)
    End If

'---���������� �����
    F_InsertFire.Show

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Sm_ShapeFormShow"
End Sub

Public Sub Sm_ExtSquareFormShow(ShpObj As Visio.Shape)
'��������� ��������� ����� ������� ������� �������

    On Error GoTo EX

'---��������� ������� �����, ����� ������ ��� ������
    F_InsertExtSquare.SetFireShape ShpObj

'---���������� �����
    F_InsertExtSquare.Show

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Sm_ExtSquareFormShow"
End Sub

