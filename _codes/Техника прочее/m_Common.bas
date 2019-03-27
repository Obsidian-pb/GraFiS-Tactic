Attribute VB_Name = "m_Common"
Option Explicit


Public Sub ShowPressureForm(ShpObj As Visio.Shape)
'    f_PressureChange.SetShp (ShpObj)
    Set f_PressureChange.currentShp = ShpObj
    f_PressureChange.Show
End Sub


Public Function IsFirstDrop(ShpObj As Visio.Shape)
'Функция проверяет вброшенали фигура впервые и если вброшена впервые добавляет строчку свойства User.InPage
    If ShpObj.CellExists("User.InPage", 0) = 0 Then
        Dim newRowIndex As Integer
        newRowIndex = ShpObj.AddNamedRow(visSectionUser, "InPage", visRowUser)
        ShpObj.CellsSRC(visSectionUser, newRowIndex, 0).Formula = 1
        ShpObj.CellsSRC(visSectionUser, newRowIndex, visUserPrompt).FormulaU = """+"""
        
        IsFirstDrop = True
    Else
        IsFirstDrop = False
    End If
End Function


