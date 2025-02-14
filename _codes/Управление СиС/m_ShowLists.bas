Attribute VB_Name = "m_ShowLists"
Option Explicit

Public ctrlOn As Boolean

'-------------------������ ��� ����������� ������� (�������������, ������ � �.�.)---------------------
Public Sub ShowUnits()
'���������� ��������� �� ����� �������������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsMainTechnics(cellVal(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---���������
    Set units = SortCol(units, "Prop.ArrivalTime", False)
    
    '��������� �������  � �������� �������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 5)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�������������"
        myArray(0, 2) = "��������"
        myArray(0, 3) = "������"
        myArray(0, 4) = "����� ��������"
        myArray(0, 5) = "������ ������"
    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.Unit", visUnitsString, "") & cellVal(shp, "Prop.Owner", visUnitsString, "")  '"�������������"
            myArray(i, 2) = cellVal(shp, "Prop.Call", visUnitsString, "") & cellVal(shp, "Prop.About", visUnitsString, "")  '"��������"
            myArray(i, 3) = cellVal(shp, "Prop.Model", visUnitsString, "")  '"������"
            myArray(i, 4) = Format(cellVal(shp, "Prop.ArrivalTime"), "DD.MM.YYYY hh:nn:ss")  '"����� ��������"
            myArray(i, 5) = cellVal(shp, "Prop.PersonnelHave")  '"������ ������"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;75 pt;100 pt;100 pt;50 pt", "ArrivedUnits", "�������"
    
End Sub

Public Sub ShowPersonnel()
'���������� ��������� �� ����� ������ ��������������� ������ � ���� ������� �������
Dim i As Integer
Dim j As Integer
Dim row As Integer
Dim shp As Visio.Shape
'Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm

Dim unitsList As Collection
Dim unitName As String
Dim callsList As String
Dim positions As Collection
Dim unitPositions As Collection
Dim pr As String
    
    
    '---�������� ������ ��������� �������� ������
    a.Refresh Application.ActivePage.Index
    Set unitsList = SortCol(a.GFSShapes, "Prop.Unit", False, visUnitsString)
    Set unitsList = GetUniqueVals(unitsList, _
                                 "Prop.Unit", , " ", " ")

    '---�������� ������ ���� ������ �������
    Set positions = FilterShapes(a.GFSShapes, "Prop.PersonnelHave;Prop.Personnel;Prop.DutyJob")
    '---���������
    Set positions = SortCol(positions, "Prop.ArrivalTime;Prop.LineTime;Prop.SetTime;Prop.FormingTime;Prop.SquareTime;Prop.FireTime", False, visDate)
    
    '��������� �������  � �������� �������
        '---������� ����� ������ ��� ����������� ������������ ������
        ReDim myArray(unitsList.Count + positions.Count, 4)
        '---������� ������ ������
        row = 0
        myArray(row, 0) = "ID"
        myArray(row, 1) = "���"
        myArray(row, 2) = "��������"
        myArray(row, 3) = "��������"
        myArray(row, 4) = "�����"
        
        '---���������� �������� ������������� � ��� ������� �� ��� ��������� �������� ��������� ������ �������
        For i = 1 To unitsList.Count
            row = row + 1
            unitName = unitsList(i)
            
            '---�������� ������ �������� ������� ������� �������������
            callsList = StrColToStr(GetUniqueVals( _
                                        FilterShapesAnd(a.GFSShapes, "Prop.PersonnelHave:;Prop.Unit:" & unitName), _
                                        "Prop.Call", , , " "), ", ")
            '---�������� �������� ������ ������� ��� ������� �������������
            Set unitPositions = FilterShapes(positions, "Prop.Unit:" & unitName)
            
            
            myArray(row, 0) = -1    '(-1 ������� ����, ��� ��� ������ ������ ��� ����� � ���������� � ��� �� �����)
            myArray(row, 1) = unitName & ":   " & callsList                      '"�������������"
            myArray(row, 2) = GetPersonnelCount(unitPositions)                   '"������ ������"
            myArray(row, 3) = CellSum(unitPositions, "Prop.Personnel")           '"�������� �/�"
            myArray(row, 4) = " "                                                '"�����"
            
            For Each shp In unitPositions
                row = row + 1

                myArray(row, 0) = shp.ID
                myArray(row, 1) = "  " & ChrW(9500) & " " & cellVal(shp, "User.IndexPers.Prompt", visUnitsString)  '"���"
                myArray(row, 2) = GetPersonnelCount(shp)                          '"������ ������"
                myArray(row, 3) = cellVal(shp, "Prop.Personnel", , " ")           '"�������� �/�"
                myArray(row, 4) = Format(pf_GetTime(shp), "DD.MM.YYYY hh:nn")     '"�����"

            Next shp
            
            '�������� ������� ��� ��������� �������� ������, ��� ������� �������
            If unitPositions.Count > 0 Then
                myArray(row, 1) = "  " & ChrW(9492) & Right(myArray(row, 1), Len(myArray(row, 1)) - 3)
            End If
        Next i
    

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;125 pt;75 pt;75 pt;100 pt", "Personnel", "������ ������"
    
End Sub

Private Function GetPersonnelCount(ByRef shps As Variant) As String
'������� ���������� �������� (��� ����� ���������)
Dim shp As Visio.Shape
Dim tmpVal As Integer
Dim tmpSum As Integer
    
    If TypeName(shps) = "Shape" Then        '���� ������
        Set shp = shps
        tmpVal = cellVal(shp, "Prop.PersonnelHave")
        If tmpVal = 0 Then
            tmpSum = tmpVal
        Else
            tmpSum = tmpVal - 1
        End If
    ElseIf TypeName(shps) = "Shapes" Or TypeName(shps) = "Collection" Then     '���� ���������
        For Each shp In shps
            tmpVal = cellVal(shp, "Prop.PersonnelHave")
            If tmpVal > 0 Then
                tmpSum = tmpSum + tmpVal - 1
            End If
        Next shp
    End If
    
    If tmpSum = 0 Then
        GetPersonnelCount = " "
    Else
        GetPersonnelCount = CStr(tmpSum)
    End If
End Function


Public Sub ShowNozzles()
'���������� ��������� �� ����� ������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsStvols(cellVal(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---���������
    Set units = SortCol(units, "Prop.SetTime", False)
    
    '��������� �������  � �������� �������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 7)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�������������"
        myArray(0, 2) = "��� ������"
        myArray(0, 3) = "��������"
        myArray(0, 4) = "����� ������"
        myArray(0, 5) = "������ ������"
        myArray(0, 6) = "������"
        myArray(0, 7) = "������������������"
    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = cellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"��� ������"
            myArray(i, 3) = cellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 4) = Format(cellVal(shp, "Prop.SetTime"), "DD.MM.YYYY hh:nn:ss")  '"����� ������"
            myArray(i, 5) = cellVal(shp, "Prop.Personnel")  '"������ ������"
            myArray(i, 6) = cellVal(shp, "Prop.UseDirection", visUnitsString, "")  '"������"
            myArray(i, 7) = cellVal(shp, "User.PodOut")  '"������������������"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt;50 pt;50 pt;50 pt", "Nozzles", "������"
    
End Sub

Public Sub ShowGDZS()
'���������� ��������� �� ����� �������� ����
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsGDZS(cellVal(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---���������
    Set units = SortCol(units, "Prop.FormingTime", False)
    
    '��������� �������  � �������� �������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 6)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�������������"
        myArray(0, 2) = "���"
        myArray(0, 3) = "��������"
        myArray(0, 4) = "����� ������������"
        myArray(0, 5) = "������ ������"
        myArray(0, 6) = "�����"
    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = cellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"���"
            myArray(i, 3) = cellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 4) = Format(cellVal(shp, "Prop.FormingTime"), "DD.MM.YYYY hh:nn:ss")  '"����� ������������"
            myArray(i, 5) = cellVal(shp, "Prop.Personnel")  '"������ ������"
            myArray(i, 6) = cellVal(shp, "Prop.AirDevice", visUnitsString, " ")  '"�����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt;50 pt;50 pt", "GDZS", "����"
    
End Sub

Public Sub ShowTimeLine()
'���������� ��������� �� ����� �������� ������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsTimeLine(shp) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
'    Set units = A.Refresh(Application.ActivePage.Index).GFSShapes
    
    '---���������
    Set units = SortCol(units, "Prop.ArrivalTime;Prop.LineTime;Prop.SetTime;Prop.FormingTime;Prop.SquareTime;Prop.FireTime", False, visDate)
    
    '��������� �������  � �������� �������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 4)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�������������"
        myArray(0, 2) = "��������"
        myArray(0, 3) = "���"
        myArray(0, 4) = "�����"

    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = cellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 3) = cellVal(shp, "User.IndexPers.Prompt", visUnitsString)  '"���"
            myArray(i, 4) = Format(pf_GetTime(shp), "DD.MM.YYYY hh:nn:ss")  '"�����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt", "TimeLine", "��������"
    
End Sub

Public Sub ShowExplication()
'���������� ������ ��������� �� ����� ��������� (�����������)
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In Application.ActivePage.Shapes
        If cellVal(shp, "User.ShapeType") = 38 Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---��������� ��������� �� ������ ���������
    Set units = SortCol(units, "Prop.LocationID", False)
    
    '��������� �������  � �������� �������� � ������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 5)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "���"
        myArray(0, 2) = "����������"
        myArray(0, 3) = "���"
        myArray(0, 4) = "�������"
        myArray(0, 5) = "���������� ����� �����"

    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.LocationID", , "")   '"���"
            myArray(i, 2) = cellVal(shp, "Prop.Use", visUnitsString, "")  '"����������"
            myArray(i, 3) = cellVal(shp, "Prop.Name", visUnitsString, "")  '"���"
            myArray(i, 4) = cellVal(shp, "Prop.visArea", visUnitsString, "")    '"�������"
            myArray(i, 5) = cellVal(shp, "Prop.OccupantCount")  '"���������� ����� �����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;200 pt;200 pt;100 pt;100 pt", "Places", "�����������"
    
End Sub

Public Sub ShowDutyList()
'���������� ������ ��������� �� ����� ����������� ���
Dim i As Integer
Dim shp As Visio.Shape
Dim persons As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set persons = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexpers.ipDutyFace) Then
            AddUniqueCollectionItem persons, shp
        End If
    Next shp
    If persons.Count = 0 Then Exit Sub

    
    '��������� �������  � �������� �������
    If persons.Count > 0 Then
        '---��������� ���������
        Set persons = SortCol(persons, "Prop.ArrivalTime", False)
        
        ReDim myArray(persons.Count, 4)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "��������� �� ������"
        myArray(0, 2) = "���������, ���"
        myArray(0, 3) = "����� ��������"
        myArray(0, 4) = "��������������"

    
        For i = 1 To persons.Count
            '---������� ��������� �������
            Set shp = persons(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.Duty", visUnitsString)                    '"��������� �� ������"
            myArray(i, 2) = cellVal(shp, "Prop.DutyJob", visUnitsString) & _
                ", " & cellVal(shp, "Prop.FIO", visUnitsString)                          '"���������, ���"
            myArray(i, 3) = cellVal(shp, "Prop.ArrivalTime", visUnitsString)             '"����� ��������"
            myArray(i, 4) = cellVal(shp, "Prop.ServiceMembership", visUnitsString)       '"��������������"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;40 pt;400 pt;100 pt", "DutyList", "����������� ���� �� ������"
    
End Sub

Public Sub ShowEvacNodes()
'���������� ������ ��������� �� ����� ����� ��������� �� ����������� �������
Dim i As Integer
Dim shp As Visio.Shape
Dim nodes As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� ������
    Set nodes = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexpers.ipEvacNode) Then
            AddOrderedNodeItem nodes, shp
        End If
    Next shp
    If nodes.Count = 0 Then Exit Sub
    
    
    '��������� �������  � �������� �������
    If nodes.Count > 0 Then
        ReDim myArray(nodes.Count, 9)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�"
        myArray(0, 2) = "�����"
        myArray(0, 3) = "���"
        myArray(0, 4) = "������"
        myArray(0, 5) = "�����"
        myArray(0, 6) = "�����"
        myArray(0, 7) = "������� �����"
        myArray(0, 8) = "����� �������"
        myArray(0, 9) = "����� �����"
        

        For i = 1 To nodes.Count
            '---������� ��������� �������
            Set shp = nodes(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.NodeNumber", visUnitsString, "")
            myArray(i, 2) = cellVal(shp, "Prop.WayClass", visUnitsString, "")
            myArray(i, 3) = cellVal(shp, "Prop.WayType", visUnitsString, "")
            myArray(i, 4) = cellVal(shp, "Prop.WayWidth", visUnitsString, "")
            myArray(i, 5) = cellVal(shp, "Prop.WayLen", visUnitsString, "")
            myArray(i, 6) = cellVal(shp, "Prop.PeopleHere", visUnitsString, "")
            myArray(i, 7) = cellVal(shp, "Prop.PeopleFlow", visUnitsString, "")
            myArray(i, 8) = cellVal(shp, "Prop.tHere", visUnitsString, "")
            myArray(i, 9) = cellVal(shp, "Prop.t_Flow", visUnitsString, "")
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;25 pt;100 pt;100 pt;50 pt;50 pt;50 pt;70 pt;70 pt;70 pt", "EvacNodes", "���� ����"

End Sub

Private Sub AddOrderedNodeItem(ByRef nodes As Collection, ByVal nodeItem As Visio.Shape)
Dim nextNode As Visio.Shape
    
    Set nextNode = FindHigherNode(nodes, nodeItem)
    '�������� ���������
    If nextNode Is Nothing Then
        nodes.Add nodeItem, CStr(nodeItem.ID)
    Else
        nodes.Add nodeItem, CStr(nodeItem.ID), CStr(nextNode.ID)
    End If
    
End Sub
Private Function FindHigherNode(ByRef nodes As Collection, ByRef nodeIn As Visio.Shape) As Visio.Shape
'���������� ������� � ��������� nodes � ������� ������ ��� � �������� (��� ������� ����� ������� ����� ����� �������� �����)
Dim node As Visio.Shape
Dim nodeNumber As Integer
Dim nodeInNumber As Integer
    
    nodeInNumber = cellVal(nodeIn, "Prop.NodeNumber")
    For Each node In nodes
        nodeNumber = cellVal(node, "Prop.NodeNumber")
        If nodeNumber > nodeInNumber Then
            Set FindHigherNode = node
            Exit Function
        End If
    Next node
End Function


'----------------------------������ ��� ������------------------------
Public Sub ShowStatists()
'���������� ������ ��������� �� ����� ���������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexpers.ipStatist) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub

    
    '��������� �������  � �������� �������
    If units.Count > 0 Then
        ReDim myArray(units.Count, 3)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "���������"
        myArray(0, 2) = "����������"
        myArray(0, 3) = "���������� �����"

    
        For i = 1 To units.Count
            '---������� ��������� �������
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.State", visUnitsString, "")   '"���������"
            myArray(i, 2) = cellVal(shp, "Prop.Info", visUnitsString, "")  '"����������"
            myArray(i, 3) = cellVal(shp, "Prop.StatistsQuatity")  '"���������� �����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;100 pt;500 pt;50 pt", "Statists", "��������"
    
End Sub

Public Sub ShowPosredniks()
'���������� ������ ��������� �� ����� �����������
Dim i As Integer
Dim shp As Visio.Shape
Dim posr As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set posr = New Collection
    For Each shp In a.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, 1002) Then
            AddUniqueCollectionItem posr, shp
        End If
    Next shp
    If posr.Count = 0 Then Exit Sub

    
    '��������� �������  � �������� �������
    If posr.Count > 0 Then
        '---��������� ���������
        Set posr = SortCol(posr, "Prop.PosrednikNumber", False)
        
        ReDim myArray(posr.Count, 4)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�����"
        myArray(0, 2) = "�� ��� ���������"
        myArray(0, 3) = "���������, ���"
        myArray(0, 4) = "����������"

    
        For i = 1 To posr.Count
            '---������� ��������� �������
            Set shp = posr(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellVal(shp, "Prop.PosrednikNumber")                            '"�����"
            myArray(i, 2) = cellVal(shp, "Prop.Target", visUnitsString, "")                 '"�� ��� ���������"
            myArray(i, 3) = cellVal(shp, "Prop.PosredinikDuty", visUnitsString) & _
                ", " & cellVal(shp, "Prop.PosredinikFIO", visUnitsString)                   '"���������, ���"
            myArray(i, 4) = cellVal(shp, "Prop.Info", visUnitsString)                       '"����������"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;25 pt;150 pt;400 pt;200 pt", "Posredniks", "����������"
    
End Sub




'----------------------------������� �������� ����� �����
Private Function pf_IsMainTechnics(ByVal a_IndexPers As Integer) As Boolean
'�������� �� ������ �������� �������
    If a_IndexPers <= 20 Or a_IndexPers = 24 Or a_IndexPers = 25 Or a_IndexPers = 26 Or a_IndexPers = 27 Or _
        a_IndexPers = 28 Or a_IndexPers = 29 Or a_IndexPers = 30 Or a_IndexPers = 31 Or a_IndexPers = 32 Or _
        a_IndexPers = 33 Or a_IndexPers = 73 Or a_IndexPers = 74 Or _
        a_IndexPers = 160 Or a_IndexPers = 161 Or a_IndexPers = 162 Or a_IndexPers = 163 Or _
        a_IndexPers = 3000 Or a_IndexPers = 3001 Or a_IndexPers = 3002 Then
        pf_IsMainTechnics = True
    Else
        pf_IsMainTechnics = False
    End If
End Function

Private Function pf_IsStvols(ByVal a_IndexPers As Integer) As Boolean
'�������� �� ������ �������� �������
    If a_IndexPers >= 34 And a_IndexPers <= 39 Or a_IndexPers = 45 Or a_IndexPers = 76 Or a_IndexPers = 77 Then
        pf_IsStvols = True
    Else
        pf_IsStvols = False
    End If
End Function

Private Function pf_IsGDZS(ByVal a_IndexPers As Integer) As Boolean
'�������� �� ������ �������� ����
    If a_IndexPers >= 46 And a_IndexPers <= 48 Or a_IndexPers = 90 Then
        pf_IsGDZS = True
    Else
        pf_IsGDZS = False
    End If
End Function

Private Function pf_IsTimeLine(ByRef a_Shape As Visio.Shape) As Boolean
'�������� �� ������ �������� ����� ���������
    If a_Shape.CellExists("Prop.ArrivalTime", 0) = True Or a_Shape.CellExists("Prop.LineTime", 0) = True _
        Or a_Shape.CellExists("Prop.SetTime", 0) = True Or a_Shape.CellExists("Prop.FormingTime", 0) = True _
        Or a_Shape.CellExists("Prop.SquareTime", 0) = True Or a_Shape.CellExists("Prop.FireTime", 0) = True _
        Then
        pf_IsTimeLine = True
    Else
        pf_IsTimeLine = False
    End If
End Function

Public Function pf_GetTime(ByRef aO_Shape As Visio.Shape, Optional ByVal default As String = "�� ����������") As String
'��������� ������� ������
On Error GoTo ex

    If aO_Shape.CellExists("Prop.ArrivalTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.ArrivalTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.LineTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.LineTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SetTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SetTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FormingTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FormingTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SquareTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SquareTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FireTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FireTime").ResultStr(visDate)
        Exit Function
    End If

Exit Function
ex:
    pf_GetTime = default
End Function
