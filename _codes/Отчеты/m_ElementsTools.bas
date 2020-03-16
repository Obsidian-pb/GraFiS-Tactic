Attribute VB_Name = "m_ElementsTools"
Option Explicit

'--------------------------Модуль для хранения инструментальных средств для работы с набором елементов-------------
Public Sub PrintElementsList(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'Печатаем списко ключей элементов споиском по строке
'Пример: PrintElementsList "fire"
Dim elem As Element
Dim coll As Collection
Dim elemShell As ElementsShell
    
    Set elemShell = New ElementsShell
    Set coll = elemShell.GetElementsCollection(searchString, searchByCallName)
    
    For Each elem In coll
        Debug.Print elem.ID
    Next elem
    
End Sub
