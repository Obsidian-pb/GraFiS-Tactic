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



Function InsertDistance(ShpObj As Visio.Shape, Optional Contex As Integer = 0)
'��������� ���������� strelki rasstoiania �� �������� ������
'---��������� ����������
Dim shpTarget As Visio.Shape
Dim shpConnection As Visio.Shape, vsO_Shape As Visio.Shape
Dim mstrConnection As Visio.Master, mstrSrelka As Visio.Master
Dim vsoCell1 As Visio.cell, vsoCell2 As Visio.cell
Dim CellFormula As String
Dim vsi_ShapeIndex As Integer
Dim lmax As Integer
Dim inppw As Boolean

vsi_ShapeIndex = 0

'    On Error GoTo EX
    InputDistanceForm.Show
    If InputDistanceForm.Flag = False Then Exit Function  '���� ��� ����� ������ - ������� �� ��������
    lmax = InputDistanceForm.lmax
    inppw = InputDistanceForm.inppw
    
    '---���������� ��� ������ � ������� ������
    For Each shpTarget In Application.ActivePage.Shapes
        If shpTarget.CellExists("User.IndexPers", 0) = True And shpTarget.CellExists("User.Version", 0) = True Then '�������� �� ������ ������� ������
'            If shpTarget.Cells("User.Version") >= CP_GrafisVersion Then  '��������� ������ ������
                vsi_ShapeIndex = shpTarget.Cells("User.IndexPers")   '���������� ������ ������ ������
                If vsi_ShapeIndex = 135 Then
                '---���������� ��������� � ��������� ������ ������ ������ � ���������
                    Set mstrConnection = ThisDocument.Masters("Distance")
                    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
                  
                    Set vsoCell1 = shpConnection.CellsU("EndX")
                    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = shpConnection.CellsU("BeginX")
                    Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                ''---��������� ����� ����� � ���� ��� �� ��������, �������
                    If shpConnection.Cells("Width").ResultIU = 0 Or shpConnection.Cells("Width").Result(visMeters) > lmax Then shpConnection.Delete
                End If

                If inppw = True And vsi_ShapeIndex = 50 Then
                     '---���������� ��������� � ��������� ������ ������ ������ � �������������
                    Set mstrConnection = ThisDocument.Masters("Distance")
                    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
                  
                    Set vsoCell1 = shpConnection.CellsU("EndX")
                    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = shpConnection.CellsU("BeginX")
                    Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    shpConnection.Cells("Prop.ArrowStyle").FormulaU = "INDEX(2,Prop.ArrowStyle.Format)"
                End If
'            End If
        End If
     Next
     
'     If Contex = 0 And vsi_ShapeIndex = 0 Then Exit Function
    
      
'---������ �����
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpObj, visSelect

Exit Function
EX:
    SaveLog Err, "InsertDistance"
End Function




