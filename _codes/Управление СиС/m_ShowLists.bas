Attribute VB_Name = "m_ShowLists"
Option Explicit


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
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsMainTechnics(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "") & CellVal(shp, "Prop.Owner", visUnitsString, "")  '"�������������"
            myArray(i, 2) = CellVal(shp, "Prop.Call", visUnitsString, "") & CellVal(shp, "Prop.About", visUnitsString, "")  '"��������"
            myArray(i, 3) = CellVal(shp, "Prop.Model", visUnitsString, "")  '"������"
            myArray(i, 4) = Format(CellVal(shp, "Prop.ArrivalTime"), "hh:mm:ss")  '"����� ��������"
            myArray(i, 5) = CellVal(shp, "Prop.PersonnelHave")  '"������ ������"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;75 pt;100 pt;100 pt;50 pt", "ArrivedUnits", "�������"
    
End Sub

Public Sub ShowNozzles()
'���������� ��������� �� ����� ������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsStvols(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"��� ������"
            myArray(i, 3) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 4) = Format(CellVal(shp, "Prop.SetTime"), "hh:mm:ss")  '"����� ������"
            myArray(i, 5) = CellVal(shp, "Prop.Personnel", "")  '"������ ������"
            myArray(i, 6) = CellVal(shp, "Prop.UseDirection", "")  '"������"
            myArray(i, 7) = CellVal(shp, "User.PodOut", "")  '"������������������"
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
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsGDZS(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"���"
            myArray(i, 3) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 4) = Format(CellVal(shp, "Prop.FormingTime"), "hh:mm:ss")  '"����� ������������"
            myArray(i, 5) = CellVal(shp, "Prop.Personnel", "")  '"������ ������"
            myArray(i, 6) = CellVal(shp, "Prop.AirDevice", visUnitsString, " ")  '"�����"
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
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"�������������"
            myArray(i, 2) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"��������"
            myArray(i, 3) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString)  '"���"
            myArray(i, 4) = Format(pf_GetTime(shp), "hh:mm:ss")  '"�����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt", "TimeLine", "��������"
    
End Sub

Public Sub ShowStatists()
'���������� ������ ��������� �� ����� ���������
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexPers.ipStatist) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.State", visUnitsString, "")   '"���������"
            myArray(i, 2) = CellVal(shp, "Prop.Info", visUnitsString, "")  '"����������"
            myArray(i, 3) = CellVal(shp, "Prop.StatistsQuatity", , "")  '"���������� �����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;100 pt;500 pt;50 pt", "Statists", "��������"
    
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
        If CellVal(shp, "User.ShapeType") = 38 Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
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
            myArray(i, 1) = CellVal(shp, "Prop.LocationID", , "")   '"���"
            myArray(i, 2) = CellVal(shp, "Prop.Use", visUnitsString, "")  '"����������"
            myArray(i, 3) = CellVal(shp, "Prop.Name", visUnitsString, "")  '"���"
            myArray(i, 4) = CellVal(shp, "Prop.visArea", visUnitsString, "")    '"�������"
            myArray(i, 5) = CellVal(shp, "Prop.OccupantCount", , "")  '"���������� ����� �����"
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;200 pt;200 pt;100 pt;100 pt", "Places", "�����������"
    
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

Private Function pf_GetTime(ByRef aO_Shape As Visio.Shape) As String
'��������� ������� ������
On Error GoTo EX

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
EX:
    pf_GetTime = "�� ����������"
End Function
