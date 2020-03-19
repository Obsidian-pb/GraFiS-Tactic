Attribute VB_Name = "m_AnalizerTools"
Option Explicit

Public analizer As InfoCollector2               '��������� ������ �����������
Private changedNames As Collection

'--------------------------������ ��� �������� ���������������� ������� ��� ������ � ������� ���������-------------
Public Function A() As InfoCollector2
'�������� ������ �� ������ �����������
    If analizer Is Nothing Then Set analizer = New InfoCollector2
    Set A = analizer
End Function
Public Sub KillA()
'���������� ������ �����������
    Set analizer = Nothing
End Sub




Public Sub FindE(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'�������� ������ ������ ��������� � ������� �� ������
'������: FindE "fire"
Dim elem As Element
Dim coll As Collection
Dim elemShell As ElementsShell
    
    Set elemShell = New ElementsShell
    Set coll = elemShell.GetElementsCollection(searchString, searchByCallName)
    
    For Each elem In coll
        Debug.Print elem.ID
    Next elem
    
End Sub



'-----------------����������� ���������� �������� � ������� �������---------------------------
Public Sub ReportShapesStringsRefresh() '(ByVal block As Boolean)
'��������� ���������! ���������� ������ ��� ���������� ����������� ��������� ������
Dim doc As Visio.Document
Dim elements As ElementsShell
Dim CallNames As String
Dim mstr As Visio.Master
Dim shp As Visio.Shape
    
    '���������� ��������� �������� �������
    Set changedNames = New Collection
    
    '������� ������ � ������� ������ ���� ���������
    Set elements = New ElementsShell
    CallNames = elements.CallNames(";")

    '���������� ��� ������� � ��������� � � ��� ���������� ��� ������.
    Set doc = ThisDocument
    For Each mstr In doc.Masters
        Debug.Print "===========: " & mstr.Name
        For Each shp In mstr.Shapes
            SetPropertyNames shp, CallNames
        Next shp
    Next mstr
    
    On Error GoTo EX
    '!��� �������� � ������ �������
    Set doc = Application.Documents("������ �������.vss")
    If Not doc Is Nothing Then
        For Each mstr In doc.Masters
            Debug.Print "===========: " & mstr.Name
            For Each shp In mstr.Shapes
                SetPropertyNames shp, CallNames
            Next shp
        Next mstr
    End If

EX:
    Set changedNames = Nothing
    Set elements = Nothing
End Sub

Private Sub SetPropertyNames(ByRef shp As Visio.Shape, ByVal properyNames As String)
'���� � ������ ������� ������ Prop.PropertyName, �������� �� ���������� ����� ������� ��������� � �������� ��������� ��� ����� ��������������.
Dim shpChild As Visio.Shape
Dim curCallName As String
Dim newCallName As String
    
    If shp.CellExists("Prop.PropertyName", 0) Then
        curCallName = CellVal(shp, "Prop.PropertyName.Value", visUnitsString)
        If Not InStr(1, properyNames, curCallName, vbTextCompare) > 0 Then
            '��������� ���������� �� ��� ����� ��� ������
            newCallName = GetExistedAnswer(curCallName)
            If newCallName = "" Then
                NewPropSelectForm.OpenForSelect properyNames, curCallName
                If NewPropSelectForm.ok Then
                    shp.Cells("Prop.PropertyName.Format").Formula = """" & properyNames & """"
                    shp.Cells("Prop.PropertyName.Value").Formula = """" & NewPropSelectForm.lbCallNames.Text & """"
                    '��������� �������� ������� ��� ��������� ��������� ����������! ������ �������� ������
                    
'                    changedNames.Add NewPropSelectForm.lbCallNames.Text, curCallName
                End If
            Else
                shp.Cells("Prop.PropertyName.Format").Formula = """" & properyNames & """"
                shp.Cells("Prop.PropertyName.Value").Formula = """" & newCallName & """"
            End If
        End If
    End If
    
    If shp.Shapes.Count > 0 Then
        For Each shpChild In shp.Shapes
            SetPropertyNames shpChild, properyNames
        Next shpChild
    End If
End Sub

Private Function GetExistedAnswer(ByVal curCallName As String) As String
Dim str As Object
    
    On Error GoTo EX

    Set str = changedNames.Item(curCallName)
Exit Function
EX:
    Set str = Nothing
End Function

