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


'Процедура задает цвет линии расстояния
Public Sub DistBuild(ShpObj As Visio.Shape)
Dim FBeg As String, FEnd As String

    FBeg = ShpObj.Cells("BegTrigger").FormulaU
    FEnd = ShpObj.Cells("EndTrigger").FormulaU

'голубой
    If InStr(1, FBeg, "ПГ") <> 0 Or InStr(1, FEnd, "ПГ") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Пирс") <> 0 Or InStr(1, FEnd, "Пирс") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Колодец") <> 0 Or InStr(1, FEnd, "Колодец") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Башня") <> 0 Or InStr(1, FEnd, "Башня") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "ПВ") <> 0 Or InStr(1, FEnd, "ПВ") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "водоисточник") <> 0 Or InStr(1, FEnd, "водоисточник") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Емкость") <> 0 Or InStr(1, FEnd, "Емкость") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
'фиолетовый
    If InStr(1, FBeg, "Здание") <> 0 And InStr(1, FEnd, "Здание") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD(RGB(112, 48, 160))"
  
End Sub
