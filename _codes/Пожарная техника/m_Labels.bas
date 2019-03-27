Attribute VB_Name = "m_Labels"
Public Sub ConnectedShapesLostCheck(ShpObj As Visio.Shape)
'Процедура проверяет, не была ли удалена фигура к которой присоединена подпись, и если была, то удаляет сам коннектор
On Error GoTo EX
    
    If Not InStr(1, ShpObj.Cells("PinX").FormulaU, "GUARD") > 0 Then
        ShpObj.Delete
    End If
Exit Sub
EX:
    'Ошибка
End Sub
