Attribute VB_Name = "archiv"
'Sub CloneSecMiscellanious(ShapeFromID As Integer, ShapeToID As Integer)
''��������� ����������� �������� ����� ��� ������ "Miscellanious"
''---��������� ����������
'Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
'Dim RowCountFrom As Integer, RowCountTo As Integer
'Dim RowNum As Integer
'
''---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
'Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1)
'Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)
'
''---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
'    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
'        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
'    Next j
'
'End Sub



'------------------------------------------���������� ����� ������-----------------------------
'Dim vO_FromShape, vO_ToShape As Visio.Shape
'Dim vi_InRowNumber, vi_OutRowNumber As Integer
'
''---���������� ���� ������ ���� ���������
'    Set vO_FromShape = Connects.FromSheet
'    Set vO_ToShape = Connects.ToSheet
'
''---���������, �������� �� ����������� ������ �������� ������
'    If vO_FromShape.CellExists("User.IndexPers", 0) = False Or _
'        vO_ToShape.CellExists("User.IndexPers", 0) = False Then Exit Sub '---��������� �������� �� ������ _
'                                                                                �������� �����
''---���������, �������� �� ����������� ������ ���������� ���
'    If f_IdentShape(vO_FromShape.Cells("User.IndexPers").Result(visNumber)) = 0 Or _
'        f_IdentShape(vO_ToShape.Cells("User.IndexPers").Result(visNumber)) = 0 Then Exit Sub
'
''---�������������� �������� � ����������� ������ - ��� ���������� ������� � ���!!!
'    '---��� From ������
'    If Left(Connects(1).FromCell.Name, 18) = ccs_InIdent Then
'        Set cpO_InShape = Connects.FromSheet
'        Set cpO_OutShape = Connects.ToSheet
'        vi_InRowNumber = Connects(1).FromCell.Row
'        vi_OutRowNumber = Connects(1).ToCell.Row
'    ElseIf Left(Connects(1).FromCell.Name, 18) = ccs_OutIdent Then
'        Set cpO_InShape = Connects.ToSheet
'        Set cpO_OutShape = Connects.FromSheet
'        vi_InRowNumber = Connects(1).ToCell.Row
'        vi_OutRowNumber = Connects(1).FromCell.Row
'    End If
'    '---��� �� ������
'    If Left(Connects(1).ToCell.Name, 18) = ccs_InIdent Then
'        Set cpO_InShape = Connects.ToSheet
'        Set cpO_OutShape = Connects.FromSheet
'        vi_InRowNumber = Connects(1).ToCell.Row
'        vi_OutRowNumber = Connects(1).FromCell.Row
'    ElseIf Left(Connects(1).ToCell.Name, 18) = ccs_OutIdent Then
'        Set cpO_InShape = Connects.FromSheet
'        Set cpO_OutShape = Connects.ToSheet
'        vi_InRowNumber = Connects(1).FromCell.Row
'        vi_OutRowNumber = Connects(1).ToCell.Row
'    End If
'
'    '---��������� ��������� ���������� ������ � �������
'       ps_LinkShapes vi_InRowNumber, vi_OutRowNumber
'
'
'    On Error Resume Next
''    Debug.Print "����������� ������: " & cpO_InShape.Name
''    Debug.Print "�������� ������: " & cpO_OutShape.Name
''    Debug.Print "������ ������: " & cpO_HoseShape.Name
''    Debug.Print Left(Connects(1).FromCell.Name, 18) & " -> " & Left(Connects(1).ToCell.Name, 18)
''    Debug.Print vO_FromShape & " -> " & vO_ToShape
'    Set cpO_InShape = Nothing
'    Set cpO_OutShape = Nothing


'----------------��� ����������� � �������� �������
''                Debug.Print cpO_InShape.Cells("User.Connects")
'                cpO_InShape.Cells("User.Connects").Formula = cpO_InShape.Cells("User.Connects") + 1
'                '---��������� ���������� ������������ �������
'                If cpO_InShape.Cells("User.Connects") > 1 Then
'                    cpO_InShape.Cells("Scratch.D1").Formula = "User.PodOut/2"
'                    cpO_InShape.Cells("Scratch.D2").Formula = "User.PodOut/2"
'                Else
'                    cpO_InShape.Cells("Scratch.D" & CStr(ai_InRowNumber + 1)).Formula = "User.PodOut"
'                End If
