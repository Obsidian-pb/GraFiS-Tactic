Attribute VB_Name = "m_Connect"
Public Sub ColConn(ShpObj As Visio.Shape)
'Процедура привязки получаемого потока !Колонки! от фигуры к которой она приклеена (если ПГ)
Dim ToShape As Integer

'---Предотвращаем появление сообщения об ошибке
On Error GoTo SubExit

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("Prop.FlowIn").FormulaU = "Sheet." & ToShape & "!Prop.Production" & ""
        ShpObj.Cells("User.WSShapeID").Formula = ToShape
    Else
        ShpObj.Cells("Prop.FlowIn").FormulaU = 0
        ShpObj.Cells("User.WSShapeID").Formula = 0
    End If

Exit Sub

SubExit:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ColConn"
End Sub

