Attribute VB_Name = "m_GE"
Option Explicit


'--------------------------------Модуль для хранения процедур работы с гидроэлеваторами-----------------
Public Sub GetTotalWaterValueForWork(ShpObj As Visio.Shape)
Dim value As Double
    value = CStr(GetConnectedHosesValue(ShpObj))
    ShpObj.Cells("User.WaterForWorkNeed").FormulaU = value ' Chr(34) & CStr(value) & Chr(34)
End Sub



Private Function GetConnectedHosesValue(ByRef shp As Visio.Shape) As Double
'Прока возвращает объем воды в присоединенных к гидроэлеватору рукавных линиях
Dim con As Connect
Dim sideShp As Visio.Shape
Dim totalHoseValue As Double
    
    On Error GoTo Tail
    
    For Each con In shp.Connects
            Set sideShp = con.ToSheet
            If sideShp.CellExists("User.IndexPers", 0) Then
                If sideShp.Cells("User.IndexPers") = 100 Then
                    totalHoseValue = totalHoseValue + _
                        sideShp.Cells("Prop.LineValue").Result(visNumber)
                End If
            End If
    Next con
    For Each con In shp.FromConnects
            Set sideShp = con.FromSheet
            If sideShp.CellExists("User.IndexPers", 0) Then
                If sideShp.Cells("User.IndexPers") = 100 Then
                    totalHoseValue = totalHoseValue + _
                        sideShp.Cells("Prop.LineValue").Result(visNumber)
                End If
            End If
    Next con

GetConnectedHosesValue = totalHoseValue
Exit Function
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "GetConnectedHosesValue"
End Function

