VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------����� ���������� ������������ ������� ������� ������------------------------------------------
Option Explicit

'---------------------���������-------------------------------------------------------------------------------------
Public ColP_Fires As New Collection
Public ColP_WStvols As New Collection


'---------------------��������--------------------------------------------------------------------------------------
Public pi_TargetPageIndex As Integer          '������� ���� �������� (0-��� �����, 1-... - ���� � ��������� �������)
Private pcpO_BeginFire As Visio.Shape         '������ ����� ������ - �������� ����� �� �������
Private pcpD_BeginDate As Date                '����� ������ ������
Private FiresColSorted As Boolean             '����������� �� ��� ��������� �������� �������
Private WStvolsColSorted As Boolean           '����������� �� ��� ��������� ������� �������

'---���������� �������� ��������
Const csi_FireIndex = 3
Const csi_ExtIndex = 4
Const csi_PodIndex = 5
Const csi_PodEffIndex = 6

'---��������� ������ �������
Const CP_GrafisVersion = 1

'����� ������������ ������ �� ��������� ����� � ���������� ������� ������ � ������������ �������
'�� ����:
'-����� �����
'-��� �������

'�� �����
'-������� ������ ��� ���������� ��������
'-����� ������ ������
'-�������������
'-����� ��������� �������




'----------------------------------------��������� ���������------------------------------------------------------------
Private Sub Class_Initialize()
'MsgBox "����� ���������������"

End Sub


'----------------------------------------������� ������� (��������� � ������� � ���)-------------------------------
Public Sub sC_ColRefresh() '(as_PageIndex As Integer)
'��������� ���������� ��������� ������
'���������� ��� �������� ����������� �� ��������� as_PageIndex �����
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer
Dim vsO_TargetPage As Visio.Page
Dim tempTime1 As Date
Dim tempTime2 As Date

    On Error GoTo EX

'---��������� ������� �������� ��� �������
    Set vsO_TargetPage = Application.ActiveDocument.Pages(pi_TargetPageIndex)

'---���������� ��� ������ � � ������ ���� ������ �������� ������� ������ ����������� �
For Each vsO_Shape In vsO_TargetPage.Shapes
    If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then '�������� �� ������ ������� ������
        If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then '��������� ������ ������
            vsi_ShapeIndex = vsO_Shape.Cells("User.IndexPers")   '���������� ������ ������ ������
            '---���������� � ������ ������ ���� ��������� ������ � ��������� ����������� ��������
            Select Case vsi_ShapeIndex
            Case Is = 64 '������� ������
                ColP_Fires.Add vsO_Shape, str(vsO_Shape.ID)
            Case Is = 34 '������ �������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 36 '�������� �������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 39 '�������� ������� �������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 70 '���� ������
                Set pcpO_BeginFire = vsO_Shape
                pcpD_BeginDate = pcpO_BeginFire.Cells("Prop.FireTime").ResultStr(visUnitsString)
            Case Is = 35 '������ ������ ������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 37 '�������� ������ ������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 45 '��������������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 22 '�������� ���������������
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "�������� ���������" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            End Select
        End If
    End If
Next vsO_Shape

'���������, ��� ��������� ��������� � ����������
FiresColSorted = False
WStvolsColSorted = False

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "sC_ColRefresh"
End Sub

'------------------------------��������� ����������� ������-------------------------------------------------
Public Sub PS_GetFireSquares(ByRef MainArray())
'��������� �������� ������������ ������ � ��������� ��� ������� � �������� ������
'������ ����� (0-���.�����; 1-�������� �������), ������ �������� ��������������� �������� �����
Dim CurShape As Visio.Shape
Dim curTime As Date
Dim curSquare As Double
Dim i As Integer

    On Error GoTo EX

'---��������� ������ � ��������� �������� �������
    sC_SortFires

'---�������������� ������ � ����������� � ����������� ����� � ���������
    If ColP_Fires.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_Fires.Count)
    End If

'---���������� ��� ������ � ���������
'    i = 0
'    For Each CurShape In ColP_Fires
'        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
'        MainArray(1, i) = CurShape.Cells("User.FireSquare")
'        i = i + 1
'    Next CurShape
    i = 0
    For Each CurShape In ColP_Fires
        If curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate)) Then
            curSquare = curSquare + CurShape.Cells("User.FireSquare")
        Else
            curSquare = CurShape.Cells("User.FireSquare")
            curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
            i = i + 1
        End If
        MainArray(0, i) = curTime
        MainArray(1, i) = curSquare
        
    Next CurShape
    
'---������� ������ ������ �� ����� ������� (���� ������� �������)
    For i = UBound(MainArray, 2) To 1 Step -1
        If Not IsEmpty(MainArray(1, i)) Then Exit For
    Next i
    ReDim Preserve MainArray(1, i)

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_GetFireSquares"
End Sub

Public Sub PS_GetExtSquares(ByRef MainArray())
'��������� �������� ������������ ������ � ��������� ��� ������� � �������� ������� ������
'������ ����� (0-���.�����; 1-�������� �������), ������ �������� ��������������� �������� �����
Dim CurShape As Visio.Shape
Dim curTime As Date
Dim curSquare As Double
Dim i As Integer

    On Error GoTo EX

'---��������� ������ � ��������� �������� �������
    sC_SortFires

'---�������������� ������ � ����������� � ����������� ����� � ���������
    If ColP_Fires.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_Fires.Count)
    End If

'---���������� ��� ������ � ���������
'    i = 0
'    For Each CurShape In ColP_Fires
'        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
'        MainArray(1, i) = CurShape.Cells("User.ExtSquare")
'        i = i + 1
'    Next CurShape
    i = 0
    For Each CurShape In ColP_Fires
        If curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate)) Then
            curSquare = curSquare + CurShape.Cells("User.ExtSquare")
        Else
            curSquare = CurShape.Cells("User.ExtSquare")
            curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
            i = i + 1
        End If
        MainArray(0, i) = curTime
        MainArray(1, i) = curSquare
        
    Next CurShape
    
'---������� ������ ������ �� ����� ������� (���� ������� �������)
    For i = UBound(MainArray, 2) To 1 Step -1
        If Not IsEmpty(MainArray(1, i)) Then Exit For
    Next i
    ReDim Preserve MainArray(1, i)

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_GetExtSquares"
End Sub

Public Sub PS_GetWStvolsPodOut(ByRef MainArray())
'��������� �������� ������������ ������ � ��������� ��� ������� � �������� ������� �������
'������ ����� (0-���.�����; 1-�������� �������), ������ �������� ��������������� �������� �����
Dim CurShape As Visio.Shape
Dim i As Integer

    On Error GoTo EX

'---��������� ������ � ��������� ������� �������
    sC_SortWStvols

'---�������������� ������ � ����������� � ����������� ����� � ���������
    If ColP_WStvols.Count = 0 Then
        Exit Sub
    Else
        i = 0
        For Each CurShape In ColP_WStvols
            If cellVal(CurShape, "Prop.UseDirection", visUnitsString) = "�������" Then
                i = i + 1
            End If
        Next CurShape
'        ReDim MainArray(1, ColP_WStvols.Count - 1)
        ReDim MainArray(1, i - 1)
    End If

'---���������� ��� ������ � ���������
    i = 0
    For Each CurShape In ColP_WStvols
        If cellVal(CurShape, "Prop.UseDirection", visUnitsString) = "�������" Then
            MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SetTime").ResultStr(visDate))
            MainArray(1, i) = CurShape.Cells("User.PodOut")
            i = i + 1
        End If
    Next CurShape

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_GetWStvolsPodOut"
End Sub

Public Sub PS_GetWStvolsEffPodOut(ByRef MainArray())
'��������� �������� ������������ ������ � ��������� ��� ������� � ����������� �������� ������� �������
'������ ����� (0-���.�����; 1-�������� �������), ������ �������� ��������������� �������� �����
Dim CurShape As Visio.Shape
Dim i As Integer

    On Error GoTo EX

'---��������� ������ � ��������� ������� �������
    sC_SortWStvols

'---�������������� ������ � ����������� � ����������� ����� � ���������
    If ColP_WStvols.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_WStvols.Count - 1)
    End If
    
'---���������� ��� ������ � ���������
    i = 0
    For Each CurShape In ColP_WStvols
        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SetTime").ResultStr(visDate))
        MainArray(1, i) = CurShape.Cells("User.PodOutEff")
        i = i + 1
    Next CurShape

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_GetWStvolsEffPodOut"
End Sub

Public Function PF_GetBeginDateTime() As Date
'������� ���������� �������� ��������������� ������� ������ ������
    PF_GetBeginDateTime = pcpD_BeginDate
End Function

Public Function PF_GetIntence(ByVal ShpPos As Integer) As Single
'������� ���������� �������� ������������� ������ ������� ������ � ������ �������� ������� ShpPos
On Error GoTo EX
    
'    If ColP_Fires.Count = 0 Then Exit Function
    
'---��������� ������ � ��������� �������� ������
    sC_SortFires
'---�������� ������
    PF_GetIntence = ColP_Fires(ShpPos).Cells("User.WaterIntense")

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetIntence"
    PF_GetIntence = 0.1
End Function

Public Function PF_GetTimeEnd(ByVal Modificator As Single, ByVal Opt As String) As Long
'������� ���������� �������� ������� ���������� � ������� ����� ������ �� ������ ���������
'����� ���������� ����� ��������� ������� ������������ ������� ������ �� �����������
'������� ���������� � ��������� ������
Dim i As Integer
    
'---��������� ������ � ��������� �������� ������
    sC_SortFires
    
    PF_GetTimeEnd = DateDiff(Opt, pcpD_BeginDate, ColP_Fires(ColP_Fires.Count).Cells("Prop.SquareTime").ResultStr(visDate)) * Modificator
End Function

Public Function PF_CheckFireBeginExist() As Boolean
'������� ���������� ����, ���� ������ ����� ��� �� �����
    PF_CheckFireBeginExist = fC_FireCheck(pi_TargetPageIndex)
End Function

Public Function GetMaxGraphSize(ByVal ShpPos) As Single
Dim maxSquare As Single
Dim exp As Single
    
    maxSquare = PF_GetMaxSquare(ShpPos)
    exp = PF_GetExpence / PF_GetIntence(ShpPos)
        
    If exp > maxSquare Then
        GetMaxGraphSize = exp
    Else
        GetMaxGraphSize = maxSquare
    End If
End Function

Public Function PF_GetMaxSquare(ByVal ShpPos) As Single
'������� ���������� �������� ������������ ������� ������
Dim i As Integer
    
On Error GoTo EX
    
    If ColP_Fires.Count = 0 Then Exit Function
    
'---��������� ������ � ��������� �������� ������
    sC_SortFires
'---�������� ������
    PF_GetMaxSquare = RUp(ColP_Fires(ShpPos).Cells("User.FireSquare"), 1.1)

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetMaxSquare"
    PF_GetMaxSquare = 320
End Function

Public Function PF_GetExpence() As Single
'������� ���������� �������� ������������� ������� ����
Dim i As Integer
Dim shp As Visio.Shape
Dim val As Double
    
On Error GoTo EX

'---�������� ������
    For Each shp In ColP_WStvols
        If cellVal(shp, "Prop.UseDirection", visUnitsString) = "�������" Then
            val = val + shp.Cells("User.PodOut").Result(visNumber)
        End If
    Next shp
    PF_GetExpence = val

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetExpence"
    PF_GetExpence = 32
End Function


'-----------------------------------��������� ���������� ��������� �� �������--------------------------------
Private Sub sC_SortFires()
'��������� ���������� ��������� � ��������� �������� ������ �� ����������� ������� �������
Dim vsi_MinTime As Long
Dim vsi_CurTime As Long
Dim vsCol_TempCol As Collection
Dim i As Integer
Dim k As Integer
Dim j As Integer

    On Error GoTo EX

'���� ��������� ��� ����������� - �������
    If FiresColSorted = True Then Exit Sub

'---��������� ����� ���������
Set vsCol_TempCol = New Collection

'---��������� ���� ���������� �� ����� ���������� ������� ������������ ���������� ��-��� � ���������
For i = 1 To ColP_Fires.Count
    vsi_MinTime = DateDiff("s", pcpD_BeginDate, ColP_Fires.Item(1).Cells("Prop.SquareTime").ResultStr(visDate))
    j = 1
    '---��������� ���� ���������� �� ����� ���������� ������� �������� ���������� ��-��� � ���������
    For k = 1 To ColP_Fires.Count
        vsi_CurTime = DateDiff("s", pcpD_BeginDate, ColP_Fires.Item(k).Cells("Prop.SquareTime").Result(visDate))
        If vsi_CurTime < vsi_MinTime Then '���� ����� �������� ��-�� ������ �������������� ������������, �� ���������� ��� �����
            vsi_MinTime = vsi_CurTime
            j = k
        End If
    Next k
    '---��������� �� ��������� ������ ��-� � ���������� ��������� �������
    vsCol_TempCol.Add ColP_Fires.Item(j), str(ColP_Fires.Item(j).ID)
    ColP_Fires.Remove j '�� �������� ��������� - ������� ���
    
Next i

'---��������� ����������� ��������� � ������������ � ���������� ���������
    Set ColP_Fires = vsCol_TempCol
    Set vsCol_TempCol = Nothing
    
'---���������, ��� ��������� ��� �������������
    FiresColSorted = True

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "sC_SortFires"
Exit Sub
End Sub


Private Sub sC_SortWStvols()
'��������� ���������� ��������� � ��������� �������� ������ �� ����������� ������� �������
Dim vsi_MinTime As Integer
Dim vsi_CurTime As Integer
Dim vsCol_TempCol As Collection
Dim i As Integer
Dim k As Integer
Dim j As Integer

    On Error GoTo EX

'���� ��������� ��� ����������� - �������
    If WStvolsColSorted = True Then Exit Sub

'---��������� ����� ���������
Set vsCol_TempCol = New Collection

'---��������� ���� ���������� �� ����� ���������� ������� ������������ ���������� ��-��� � ���������
For i = 1 To ColP_WStvols.Count
    vsi_MinTime = DateDiff("s", pcpD_BeginDate, ColP_WStvols.Item(1).Cells("Prop.SetTime").ResultStr(visDate))
    j = 1
    '---��������� ���� ���������� �� ����� ���������� ������� �������� ���������� ��-��� � ���������
    For k = 1 To ColP_WStvols.Count
        vsi_CurTime = DateDiff("s", pcpD_BeginDate, ColP_WStvols.Item(k).Cells("Prop.SetTime").ResultStr(visDate))
        If vsi_CurTime < vsi_MinTime Then '���� ����� �������� ��-�� ������ �������������� ������������, �� ���������� ��� �����
            vsi_MinTime = vsi_CurTime
            j = k
        End If
    Next k
    '---��������� �� ��������� ������ ��-� � ���������� ��������� �������
    vsCol_TempCol.Add ColP_WStvols.Item(j), str(ColP_WStvols.Item(j).ID)
    ColP_WStvols.Remove j '�� �������� ��������� - ������� ���
    
Next i

'---��������� ����������� ��������� � ������������ � ���������� ���������
    Set ColP_WStvols = vsCol_TempCol
    Set vsCol_TempCol = Nothing
    
'---���������, ��� ��������� ��� �������������
    WStvolsColSorted = True

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "sC_SortWStvols"
End Sub


'-------------------------------------------------��������� �������-----------------------------------------------------
Private Function fC_MaxValue(afi_ColIndex As Byte, afs_CellName As String) As Variant
'������� ���������� ������������ �������� ��������� ������ � ��������� ���������
Dim vfO_TempCol As Collection
Dim vfO_Shape As Visio.Shape
Dim vfs_MaxValue As Single

    Set vfO_TempCol = New Collection
    vfs_MaxValue = 0

    Select Case afi_ColIndex
        Case Is = 1
            Set vfO_TempCol = ColP_Fires
        Case Is = 2
            Set vfO_TempCol = ColP_WStvols
    End Select

    For Each vfO_Shape In vfO_TempCol
        If vfO_Shape.Cells(afs_CellName).Result(visNumber) > vfs_MaxValue Then _
            vfs_MaxValue = vfO_Shape.Cells(afs_CellName).Result(visNumber)
    Next vfO_Shape

Set vfO_TempCol = Nothing
fC_MaxValue = vfs_MaxValue
End Function


Private Function fC_FireCheck(as_PageIndex As Integer) As Boolean
'������� �������� ������� ������ ����� ������ �� �����
Dim vfO_Page As Visio.Page
Dim vfO_Shape As Visio.Shape

fC_FireCheck = False

    Set vfO_Page = Visio.Application.ActiveDocument.Pages(as_PageIndex)
    
    For Each vfO_Shape In vfO_Page.Shapes
        If vfO_Shape.CellExists("User.IndexPers", 0) = True Then
            Debug.Print vfO_Shape.Name & " - " & vfO_Shape.Cells("User.IndexPers").Result(visNumber)
            If vfO_Shape.Cells("User.IndexPers").Result(visNumber) = 70 Then
                fC_FireCheck = True
                Set vfO_Page = Nothing
                Exit Function
            End If
        End If
    Next vfO_Shape
    
    Set vfO_Page = Nothing
    
fC_FireCheck = False
End Function






