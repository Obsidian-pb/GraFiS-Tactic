Attribute VB_Name = "m_Tools"



Public Function CD_MasterExists(masterName As String) As Boolean
'������� �������� ������� ������� � �������� ���������
Dim i As Integer

For i = 1 To Application.ActiveDocument.Masters.count
    If Application.ActiveDocument.Masters(i).Name = masterName Then
        CD_MasterExists = True
        Exit Function
    End If
Next i

CD_MasterExists = False

End Function

Public Sub MasterImportSub(masterName As String)
'��������� ������� ������� � ������������ � ������
Dim mstr As Visio.Master

    If Not CD_MasterExists(masterName) Then
        Set mstr = ThisDocument.Masters(masterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

End Sub



'-----------------����� ��� �����-------------------------------------------------
Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ - �����, � ��������� ������ - ����
    
'---���������, �������� �� ������ ������� �����������
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---���������, �������� �� ������ ������� �����
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function

Public Function PFB_isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ - ������� �����, � ��������� ������ - ����
    
'---���������, �������� �� ������ ������� �����������
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isDoor = False
        Exit Function
    End If

'---���������, �������� �� ������ ������� ����� ��� �����
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Or aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25) Then
        PFB_isDoor = True
        Exit Function
    End If
PFB_isDoor = False
End Function

Public Function PFB_isWindow(ByRef aO_Shape As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ - ����, � ��������� ������ - ����
    
'---���������, �������� �� ������ ������� �����������
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWindow = False
        Exit Function
    End If

'---���������, �������� �� ������ ������� ����
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        aO_Shape.Cells("User.ShapeType").Result(visNumber) = 45 Then
        PFB_isWindow = True
        Exit Function
    End If
PFB_isWindow = False
End Function

'--------------------------------������ �� ������-------------------------------------
Public Function GetLayerNumber(ByRef layerName As String) As Integer
Dim layer As Visio.layer

    For Each layer In Application.ActivePage.Layers
        If layer.Name = layerName Then
            GetLayerNumber = layer.Index - 1
            Exit Function
        End If
    Next layer
    
    Set layer = Application.ActivePage.Layers.Add(layerName)
    GetLayerNumber = layer.Index - 1
End Function

'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub

'---------------------------------------��������� ������� � �����--------------------------------------------------
Public Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'������� ���������� ���� ������������ ������������� ��������
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

'Set Shape = Nothing
End Function

Public Sub ClearLayer(ByVal layerName As String)
'������� ������ ���������� ����
    On Error Resume Next
    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, layerName)
    vsoSelection.Delete
End Sub

Public Function ShapeIsLine(ByRef shp As Visio.Shape) As Boolean
'������� ���������� ������, ���� ���������� ������ - ������� ������ �����, ���� - ���� �����
Dim isLine As Boolean
Dim isStrait As Boolean
    
    ShapeIsLine = False
    
    On Error GoTo EX
    
    If shp.RowCount(visSectionFirstComponent) <> 3 Then Exit Function       '����� � ������ ��������� ������ ��� ������ ����
    If shp.RowType(visSectionFirstComponent, 2) <> 139 Then Exit Function   '139 - LineTo
    
ShapeIsLine = True
Exit Function

EX:
    ShapeIsLine = False
End Function

'--------------------------------------������ � ���������-------------------------------------------------------------
Public Function GetCommandBarTool(ByRef cbr As Office.CommandBar, ByVal toolID As Integer) As Office.CommandBarControl
'������� ���������� ������ � ��������� ID ��������� ������ ������������
Dim btnTool As Office.CommandBarControl
    For Each btnTool In cbr.Controls
        If btnTool.ID = toolID Then
            Set GetCommandBarTool = btnTool
            Exit Function
        End If
    Next btnTool
Set GetCommandBarTool = Nothing
End Function
