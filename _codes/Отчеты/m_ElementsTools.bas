Attribute VB_Name = "m_ElementsTools"
Option Explicit

Public analizer As InfoCollector2               'Публичный объект анализатора


'--------------------------Модуль для хранения инструментальных средств для работы с набором елементов-------------
Public Sub FindE(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'Печатаем списко ключей элементов споиском по строке
'Пример: FindE "fire"
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
'Получаем ссылку на объект анализатора
    If analizer Is Nothing Then Set analizer = New InfoCollector2
    Set A = analizer
End Function
Public Sub KillA()
'Уничтожаем объект анализатора
    Set analizer = Nothing
End Sub
