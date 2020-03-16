Attribute VB_Name = "m_ElementsTools"
Option Explicit

'--------------------------������ ��� �������� ���������������� ������� ��� ������ � ������� ���������-------------
Public Sub PrintElementsList(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'�������� ������ ������ ��������� �������� �� ������
'������: PrintElementsList "fire"
Dim elem As Element
Dim coll As Collection
Dim elemShell As ElementsShell
    
    Set elemShell = New ElementsShell
    Set coll = elemShell.GetElementsCollection(searchString, searchByCallName)
    
    For Each elem In coll
        Debug.Print elem.ID
    Next elem
    
End Sub
