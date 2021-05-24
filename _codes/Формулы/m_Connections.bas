Attribute VB_Name = "m_Connections"
Option Explicit



Public Sub TurnIntoFormulaConnection(ByRef Connects As IVConnects)
'Процедура обращения в коннектор формул (при соединении)
Dim cnct As Visio.Connect
Dim shp As Visio.Shape
Dim rowI As Integer
    
    On Error GoTo EndSub
    
    'Опредяелмя фигуру коннектор
    Set shp = Connects(1).FromSheet
    
    '---проверяем не является ли фигуры фигурой коннектора уже
    If IsGFSShapeWithIP(shp, 501, True) Then Exit Sub
    
    '---проверяем, имеет ли фигура ДВЕ точки соединения
    If shp.Connects.Count <> 2 Then Exit Sub

    '---Проверяем Является ли фигура линией
    If shp.AreaIU > 0 Then Exit Sub
    
    '--Проверяем обе ли соединенных фигур являются фигурами формул
    If IsGFSShapeWithIP(shp.Connects(1).ToSheet, 500, True) And IsGFSShapeWithIP(shp.Connects(2).ToSheet, 500, True) Then
        '---Основная процедура обращения
        f_LinkToCell2.showForm shp.Connects(1).ToSheet, shp.Connects(2).ToSheet, shp
        
        'Добавляем коннектору свойства
        shp.AddNamedRow visSectionUser, "IndexPers", visTagDefault
        SetCellVal shp, "User.IndexPers", "501"
    End If
    
Exit Sub
EndSub:
    SaveLog Err, "TurnIntoFormulaConnection"
End Sub
