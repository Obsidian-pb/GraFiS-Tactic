Attribute VB_Name = "t_COllections"
Option Explicit


Public Function SortCommands(ByVal coms As Collection, Optional ByVal desc As Boolean = False) As Collection
'Функция возвращает отсортированную коллекцию фигур. Фигуры сортируются по значению в ячейке sortCellName - чем больше, тем выше
'desc: True - от большего к меньшему; False - от меньшего к большему
Dim i As Integer
Dim com As c_SimpleDescription
Dim tmpColl As Collection

    
    Set tmpColl = New Collection
    
    Do While coms.Count > 1
        
'        If desc Then
'            Set com = GetMaxCom(shps)
'        Else
            Set com = GetMinCom(coms)
'        End If
        
        AddUniqueCollectionItem tmpColl, com
        RemoveFromCollection coms, com
        
        i = i + 1
        If i > 100 Then Exit Do
    Loop
    
    Set com = coms(1)
    AddUniqueCollectionItem tmpColl, com
'    Debug.Print tmpshp.ID & " " & cellVal(tmpshp, sortCellName)
    
    Set SortCommands = tmpColl
End Function
Public Function GetMinCom(ByRef col As Collection) As c_SimpleDescription
'Возвращает команду с минимальным временем отдачи
Dim i As Integer
Dim com1 As c_SimpleDescription
Dim com2 As c_SimpleDescription
Dim com1Time As Double
Dim com2Time As Double
    
    Set com1 = col(1)
    com1Time = com1.time
    For i = 1 To col.Count
        Set com2 = col(i)
        com2Time = com2.time
        
        If com2Time < com1Time Then
            Set com1 = com2
            com1Time = com2Time
        End If
    Next i
'    Debug.Print com1.time & " " & com1.text
    
Set GetMinCom = com1
End Function
