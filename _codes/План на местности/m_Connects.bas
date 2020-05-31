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
Dim ShpSubj As Visio.Shape

    FBeg = ShpObj.Cells("BegTrigger").FormulaU
    FEnd = ShpObj.Cells("EndTrigger").FormulaU

'�������
    If InStr(1, FBeg, "��") <> 0 Or InStr(1, FEnd, "��") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "����") <> 0 Or InStr(1, FEnd, "����") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "�������") <> 0 Or InStr(1, FEnd, "�������") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "�����") <> 0 Or InStr(1, FEnd, "�����") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "��") <> 0 Or InStr(1, FEnd, "��") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "������������") <> 0 Or InStr(1, FEnd, "������������") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "�������") <> 0 Or InStr(1, FEnd, "�������") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If

'����������
    If InStr(1, FBeg, "������") <> 0 And InStr(1, FEnd, "������") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD(RGB(112, 48, 160))"
        GoTo SubjDistance
    End If
    
Exit Sub
SubjDistance:
    ''---�������� �������� �������� ����� ��� ������������ ����� � ��������� �����������
    Set ShpSubj = ActivePage.Shapes(Replace(Replace(FBeg, "_XFTRIGGER(", ""), "!EventXFMod)", ""))
    If ShpSubj.CellExists("User.Distance", 0) = False Then ShpSubj.AddNamedRow visSectionUser, "Distance", 0
    ShpSubj.Cells("User.Distance").FormulaU = "Sheet." & ShpObj.ID & "!User.Length"
    
    Set ShpSubj = ActivePage.Shapes(Replace(Replace(FEnd, "_XFTRIGGER(", ""), "!EventXFMod)", ""))
    If ShpSubj.CellExists("User.Distance", 0) = False Then ShpSubj.AddNamedRow visSectionUser, "Distance", 0
    ShpSubj.Cells("User.Distance").FormulaU = "Sheet." & ShpObj.ID & "!User.Length"
  
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
           '��������� �� ������� �� ���������� �� ������ ���
            If shpTarget.CellExists("User.Distance", 0) = False Then shpTarget.AddNamedRow visSectionUser, "Distance", 0
               If InStr(1, shpTarget.Cells("User.Distance").FormulaU, "!User.Length") = 0 Then
                vsi_ShapeIndex = shpTarget.Cells("User.IndexPers")   '���������� ������ ������ ������
                If vsi_ShapeIndex = 135 Then
                '---���������� ��������� � ��������� ������ ������ ������ � ���������
                    Set mstrConnection = ThisDocument.Masters("Distance2")
                    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
                  
                    Set vsoCell1 = shpConnection.CellsU("EndX")
                    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = shpConnection.CellsU("BeginX")
                    Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                ''---��������� ����� ����� � ���� ��� �� ��������, �������
DoEvents '��� ����� ����� ����� ������������� � �����������
                    DistBuild shpConnection
                    If shpConnection.Cells("User.Length").Result(visMeters) = 0 Or _
                    shpConnection.Cells("User.Length").Result(visMeters) > lmax Then shpConnection.Delete
                End If

                If inppw = True And (vsi_ShapeIndex = 50 Or vsi_ShapeIndex = 51 Or vsi_ShapeIndex = 53 _
                   Or vsi_ShapeIndex = 54 Or vsi_ShapeIndex = 55 Or vsi_ShapeIndex = 56 Or vsi_ShapeIndex = 240 Or vsi_ShapeIndex = 190) Then
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
                    shpConnection.Cells("Controls.ArrowPos").FormulaU = "IF(Width<6 m,Width*0.7*ThePage!User.GFS_Aspect,5 m*ThePage!User.GFS_Aspect)"
                    shpConnection.Cells("Controls.TxtPos").FormulaU = "IF(STRSAME(Prop.ArrowStyle,""���������������� �����������""),IF(Width<6 m,Width*0.5*ThePage!User.GFS_Aspect,2.5 m*ThePage!User.GFS_Aspect),Width*0.5*ThePage!User.GFS_Aspect)"
                    shpConnection.Cells("Controls.Ident").FormulaU = "IF(Width<6 m,Width*0.3*ThePage!User.GFS_Aspect,2 m *ThePage!User.GFS_Aspect)"
                    DistBuild shpConnection
                End If
               End If
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




