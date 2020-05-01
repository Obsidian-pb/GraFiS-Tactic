Attribute VB_Name = "Vichislenia"
Option Explicit

Sub SquareSet(ShpObj As Visio.Shape)
'Процедура присвоения текстовому полю выделенной фигуры значения площади фигуры
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 'переводим из квадратных дюймов в квадратные метры
ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub


Sub s_SetFireTime(ShpObj As Visio.Shape, Optional showDoCmd As Boolean = True)
'Процедура присвоения ячейке документа User.FireTime значения времени указанного при вбрасывании фигуры "Очаг"
Dim vD_CurDateTime As Double

On Error Resume Next

'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Присваиваем значению времени возникновения пожара текущее значение
            vD_CurDateTime = Now()
            ShpObj.Cells("Prop.FireTime").FormulaU = _
                "DATETIME(" & str(vD_CurDateTime) & ")"
        
        '---Показываем окно свойств фигуры
            If showDoCmd Then Application.DoCmd (1312)
            
        '---Если в Шэйп-листе документа отсутствует строка "User.FireTime", создаем её
            If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
                Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
            End If
            
        '---Переносим новые данные из шейп личста фигуры в шейп лист документа
            Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
                "DATETIME(" & str(CDbl(ShpObj.Cells("Prop.FireTime").Result(visDate))) & ")"
    Else
        '---Показываем окно свойств фигуры
            If showDoCmd Then Application.DoCmd (1312)
    End If

End Sub
