Attribute VB_Name = "m_Connects"
Option Explicit
'---------------------------------Модуль для хранения процедур линковки данных-----------------------------

Public Sub Conn(ShpObj As Visio.Shape)
'Процедура привязки отображаемого значения в подписи к фигуре к котрой она приклеена
Dim ToShape As Long

'---Предотвращаем появление сообщения об ошибке
On Error Resume Next

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = "Sheet." & ToShape & "!Prop.Street" & ""
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = 0
    End If

'---Перемещаем фигуру поверх прочих (На передний план)
    ShpObj.BringToFront
End Sub
