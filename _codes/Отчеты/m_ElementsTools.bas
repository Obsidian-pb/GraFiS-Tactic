Attribute VB_Name = "m_ElementsTools"
Option Explicit

Public analizer As InfoCollector2               '��������� ������ �����������


'--------------------------������ ��� �������� ���������������� ������� ��� ������ � ������� ���������-------------
Public Sub FindE(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'�������� ������ ������ ��������� �������� �� ������
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

Public Function A() As InfoCollector2
'�������� ������ �� ������ �����������
    If analizer Is Nothing Then Set analizer = New InfoCollector2
    Set A = analizer
End Function
Public Sub KillA()
'���������� ������ �����������
    Set analizer = Nothing
End Sub
