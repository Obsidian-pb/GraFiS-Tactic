Attribute VB_Name = "Archiv"
''----------------------------------������ ����������� ������ �������� ��� � ������� ��� ������� �����--------------------------------
'
'
''Public Sub DDD()
''MsgBox DateDiff("n", "29.10.2012 0:45:00", "29.10.2012 3:50:00")
''End Sub
'
'
'Public Sub sP_WaterSiSReportDrop(ShpObj As Visio.Shape)
'ShpObj.Delete
''��������� ����������� ������ �������� ��� � ������� ��� ������� �����
''---��������� ����������
'Dim vpVS_DocShape As Visio.Shape
'Dim vps_pth As String
'Dim vpO_ExelApp As Excel.Application
'Dim vpWB_ExelWB As Excel.Workbook
'
''---�������� ����� �������� ��������� � �� ��� ��������� ������� ����� ��������� ������ � �������
'vps_pth = ThisDocument.path & "XLS\Reports.xls"
'
''---��������� ������
'Set vpO_ExelApp = New Excel.Application
'Set vpWB_ExelWB = vpO_ExelApp.Workbooks.Open(vps_pth)
''vpO_ExelApp.Visible = True
'
''---������ ��������� � ��������
'    vpWB_ExelWB.Worksheets(1).Range("D2").Value = fps_TotalFireSquare
'    vpWB_ExelWB.Worksheets(1).Range("D3").Value = fps_TotalExtSquare
'    vpWB_ExelWB.Worksheets(1).Range("D4").Value = fps_WaterIntense
'    vpWB_ExelWB.Worksheets(1).Range("D6").Value = fps_ShapeSum(34) + fps_ShapeSum(36) + fps_ShapeSum(39)
'    vpWB_ExelWB.Worksheets(1).Range("D7").Value = fpi_ShapeCount(34) + fpi_ShapeCount(36) + fpi_ShapeCount(39)
'    vpWB_ExelWB.Worksheets(1).Range("C8").Value = fpi_PersonnelNeedSum + fpi_HosePersonnel '���-�� ������� ������� ������������
'    vpWB_ExelWB.Worksheets(1).Range("D8").Value = fpi_PersonnelHaveSum '���-�� ������� ������� ����������
'    vpWB_ExelWB.Worksheets(1).Range("D9").Value = fpi_ShapeCount(1) + fpi_ShapeCount(2)
'    vpWB_ExelWB.Worksheets(1).Range("D10").Value = fpi_ShapeCount(72) + fpi_ShapeCount(88) '���-�� ����������� ������������� �� ��
'    vpWB_ExelWB.Worksheets(1).Range("D12").Value = fpi_ShapeCount(46) + fpi_ShapeCount(90) '���-�� ������� ����
'    vpWB_ExelWB.Worksheets(1).Range("D13").Value = fpi_ShapeCount(9) + fpi_ShapeCount(10) + _
'        fpi_ShapeCount(11) + fpi_ShapeCount(12) + fpi_ShapeCount(13) + fpi_ShapeCount(14)   '���-�� ����������� �������� ����������
'    vpWB_ExelWB.Worksheets(1).Range("D14").Value = fpi_ShapeCount(3) + fpi_ShapeCount(4) + _
'        fpi_ShapeCount(5) + fpi_ShapeCount(6) + fpi_ShapeCount(7) + fpi_ShapeCount(8) + _
'        fpi_ShapeCount(15) + fpi_ShapeCount(16) + _
'        fpi_ShapeCount(17) + fpi_ShapeCount(18) + fpi_ShapeCount(19) + fpi_ShapeCount(20)   '���-�� ����������� �����������
'
'
'
'
'
''---���������� ���������� ��������
'Visio.Application.ActiveWindow.Page.InsertFromFile vps_pth, visInsertAsEmbed
'
''---��������� �������� � ����������
'vpWB_ExelWB.Close (False)
'vpO_ExelApp.Quit
'
'
'End Sub
'
''-----------------------------------------������� �������� ���������� �����������----------------------------
'Private Function fps_TotalFireSquare() As Single
''������� �������� ����� ������� ������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vps_Temp As Single
'Dim vpd_Max As Date
'
''---���������� ��� ������ � � ������ ���� ������ �������� �������� ������ ��������� �
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = 64 Then
'            If vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime") >= vpd_Max Then '---��������� �������� �� ������ ������ ��������� �� �������
'                vps_Temp = vpO_Shape.Cells("User.FireSquare")
'                vpd_Max = vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime")
'            End If
'        End If
'    End If
'Next vpO_Shape
'
'
'fps_TotalFireSquare = vps_Temp
'End Function
'
'Private Function fps_TotalExtSquare() As Single
''������� �������� ����� ������� �������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vps_Temp As Single
'Dim vpd_Max As Date
'
''---���������� ��� ������ � � ������ ���� ������ �������� �������� ������ ��������� ������� � �������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = 64 Then
'            If vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime") >= vpd_Max Then '---��������� �������� �� ������ ������ ��������� �� �������
'                vps_Temp = vpO_Shape.Cells("User.ExtSquare")
'                vpd_Max = vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime")
'            End If
'        End If
'    End If
'Next vpO_Shape
'
'
'fps_TotalExtSquare = vps_Temp
'End Function
'
'
'Private Function fps_WaterIntense() As Single
''������� �������� �������������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vps_Temp As Single
'Dim vpd_Max As Date
'
''---���������� ��� ������ � � ������ ���� ������ �������� �������� ������ ������ � �������������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = 64 Then
'            If vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime") >= vpd_Max Then '---��������� �������� �� ������ ������ ��������� �� �������
'                vps_Temp = vpO_Shape.Cells("User.WaterIntense")
'                vpd_Max = vpO_Shape.Cells("Prop.SquareDate") + vpO_Shape.Cells("Prop.SquareTime")
'            End If
'        End If
'    End If
'Next vpO_Shape
'
'
'fps_WaterIntense = vps_Temp
'End Function
'
'
'Private Function fpi_ShapeCount(vf_IndPers As Integer) As Integer
''������� �������� ����������� ����� � ��������� IndexPers
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vpi_Temp As Integer
'
''---���������� ��� ������ � � ������ ���� ������ ����� ��������� IndexPers ����������� ����������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = vf_IndPers Then
'            vpi_Temp = vpi_Temp + 1
'        End If
'    End If
'Next vpO_Shape
'
'fpi_ShapeCount = vpi_Temp
'End Function
'
'
'Private Function fps_ShapeSum(vf_IndPers As Integer) As Single
''������� �������� ���� ��������� ������� ����� � ��������� IndexPers
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vps_Temp As Single
'
''---���������� ��� ������ � � ������ ���� ������ ����� ��������� IndexPers ��������� ��������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = vf_IndPers Then
'            vps_Temp = vps_Temp + vpO_Shape.Cells("User.PodOut")
'        End If
'    End If
'Next vpO_Shape
'
'fps_ShapeSum = vps_Temp
'End Function
'
'
'Private Function fpi_PersonnelNeedSum() As Integer
''������� �������� ���� ���������� ������� �������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vpi_Temp As Integer
'
''---���������� ��� ������ � � ������ ���� ������ ����� �������� "Personnel" ��������� ��������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("Prop.Personnel", 0) = True Then
'            vpi_Temp = vpi_Temp + vpO_Shape.Cells("Prop.Personnel")
'    End If
'Next vpO_Shape
'
'fpi_PersonnelNeedSum = vpi_Temp
'End Function
'
'
'Private Function fpi_PersonnelHaveSum() As Integer
''������� �������� ���� ���������� ������� �������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vpi_Temp As Integer
'
''---���������� ��� ������ � � ������ ���� ������ ����� �������� "PersonnelHave" ��������� ��������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("Prop.PersonnelHave", 0) = True Then
'            vpi_Temp = vpi_Temp + vpO_Shape.Cells("Prop.PersonnelHave") - 1 '�� ������� ���������
'    End If
'Next vpO_Shape
'
'fpi_PersonnelHaveSum = vpi_Temp
'End Function
'
'Private Function fpi_HosePersonnel() As Integer
''������� �������� ���������� ������� ������� ������������ ��� �������� �������� ������
''---��������� ����������
'Dim vpO_Shape As Visio.Shape
'Dim vpi_TotalHosesLenight As Integer
'Dim vpi_PersonnelNeed As Integer
'
''---���������� ��� ������ � � ������ ���� ������ �������� �������� ������, ������������ ����� ������ �������
'For Each vpO_Shape In Application.ActivePage.Shapes
'    If vpO_Shape.CellExists("User.IndexPers", 0) = True Then
'        If vpO_Shape.Cells("User.IndexPers") = 100 Then
'                vpi_TotalHosesLenight = vpi_TotalHosesLenight + vpO_Shape.Cells("Prop.LineLenightHose")
'        End If
'    End If
'Next vpO_Shape
'
'vpi_PersonnelNeed = Abs(vpi_TotalHosesLenight / 100)
'
'fpi_HosePersonnel = vpi_PersonnelNeed
'End Function
'
'
''Public Sub Proverka(ShpObj As Visio.Shape)
''ShpObj.Cells("User.Row1").FormulaU = "123"
''End Sub
'
'
''Private Sub sC_SortMass()
'''Const Temp = ColP_Fires.Count
''Dim vsM_FiresArr() As Integer
''Dim tmp As Integer
''Dim i, k, j As Integer
''Dim minVal As Integer
''
''ReDim vsM_FiresArr(ColP_Fires.Count, 3)
''
''For i = 1 To ColP_Fires.Count
''    vsM_FiresArr(i, 0) = Minute(ColP_Fires.Item(i).Cells("Prop.SquareTime") - pcpD_BeginDate)
''Next i
''
''For i = 1 To UBound(vsM_FiresArr())
''    minVal = vsM_FiresArr(i, 0)
''    For k = i To UBound(vsM_FiresArr())
''        If vsM_FiresArr(k, 0) < vsM_FiresArr(i, 0) And vsM_FiresArr(k, 0) < minVal Then
''            j = k
''            minVal = vsM_FiresArr(k, 0)
''        End If
''    Next k
''    tmp = vsM_FiresArr(i, 0)
''    vsM_FiresArr(i, 0) = vsM_FiresArr(j, 0)
''    vsM_FiresArr(j, 0) = tmp
''Next i
''
''For i = 1 To UBound(vsM_FiresArr())
''    MsgBox vsM_FiresArr(i, 0)
''Next i
''
''End Sub
'
'
'Private Sub sC_SortMass()
''Const Temp = ColP_Fires.Count
'Dim vsM_FiresArr() As Integer
'Dim tmp As Integer
'Dim i, j, k As Integer
'
'ReDim vsM_FiresArr(ColP_Fires.Count, 3)
'
'For i = 1 To ColP_Fires.Count
'    vsM_FiresArr(i, 0) = Minute(ColP_Fires.Item(i).Cells("Prop.SquareTime") - pcpD_BeginDate)
'Next i
'
'For i = 1 To UBound(vsM_FiresArr())
'    For k = i To UBound(vsM_FiresArr())
'        If vsM_FiresArr(i, 0) > vsM_FiresArr(k, 0) Then
'            tmp = vsM_FiresArr(i, 0)
'            vsM_FiresArr(i, 0) = vsM_FiresArr(k, 0)
'            vsM_FiresArr(k, 0) = tmp
'        End If
'    Next k
'Next i
'
'For i = 1 To UBound(vsM_FiresArr())
'    MsgBox vsM_FiresArr(i, 0)
'Next i
'
'End Sub
'
'
'
''For i = 1 To ColP_Fires.Count
''        '---�������� � �������� ������� ������
''        '---��������� ����� ������
''        vsO_FireLine.AddRow visSectionFirstComponent, i + 1, visTagLineTo
''
''        Set vsO_ShapeElement = ColP_Fires.Item(i)
''
''        '---���������� ����� ������� ������� ������
''        vsS_Square = vsO_ShapeElement.Cells("User.FireSquare").Result(visNumber)
''        vsS_ResultRel = Round((vsS_Square / vsS_FirePrice) * vsS_FireStepPrice, 4)
''        vss_Formula = "(Sheet." & asi_ShapeID & "!Height *" & Str(vsS_ResultRel) & _
''            ")*Sheet." & asi_ShapeID & "!User.YScale"      '!!!=(Sheet.132!Height*0.0541)*Sheet.132!User.Scale
''
''        vsO_FireLine.CellsSRC(visSectionFirstComponent, i + 1, 1).FormulaU = vss_Formula
''
''        vsS_Time = Minute(vsO_ShapeElement.Cells("Prop.SquareTime") - pcpD_BeginDate)
''
''        vsS_ResultRel = Round((vsS_Time / vsS_TimePrice) * vsS_TimeStepPrice, 4)
''        vss_Formula = "(Sheet." & asi_ShapeID & "!Width *" & Str(vsS_ResultRel) & _
''            ")*Sheet." & asi_ShapeID & "!User.XScale" '!!!=(Sheet.132!Height*0.0541)*Sheet.132!User.Scale
''
''        vsO_FireLine.CellsSRC(visSectionFirstComponent, i + 1, 0).FormulaU = vss_Formula
''
''        '---�������� � �������� ������� ������
''        '---��������� ����� ������
''        vsO_ExtLine.AddRow visSectionFirstComponent, i + 1, visTagLineTo
''
''        Set vsO_ShapeElement = ColP_Fires.Item(i)
''
''        '---���������� ����� ������� ������� ������
''        vsS_Square = vsO_ShapeElement.Cells("User.ExtSquare").Result(visNumber)
''        vsS_ResultRel = Round((vsS_Square / vsS_FirePrice) * vsS_FireStepPrice, 4)
''        vss_Formula = "(Sheet." & asi_ShapeID & "!Height *" & Str(vsS_ResultRel) & _
''            ")*Sheet." & asi_ShapeID & "!User.YScale"      '!!!=(Sheet.132!Height*0.0541)*Sheet.132!User.Scale
''
''        vsO_ExtLine.CellsSRC(visSectionFirstComponent, i + 1, 1).FormulaU = vss_Formula
''
''        vsS_Time = Minute(vsO_ShapeElement.Cells("Prop.SquareTime") - pcpD_BeginDate)
''
''        vsS_ResultRel = Round((vsS_Time / vsS_TimePrice) * vsS_TimeStepPrice, 4)
''        vss_Formula = "(Sheet." & asi_ShapeID & "!Width *" & Str(vsS_ResultRel) & _
''            ")*Sheet." & asi_ShapeID & "!User.XScale" '!!!=(Sheet.132!Height*0.0541)*Sheet.132!User.Scale
''
''        vsO_ExtLine.CellsSRC(visSectionFirstComponent, i + 1, 0).FormulaU = vss_Formula
''
''    Next i
'�������� ��� GitHub desctop
