Attribute VB_Name = "Tools"
Option Explicit


'-----------------------------------------Процедуры работы с фигурами----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'Процедура устанавливает новое значение для всех выбранных фигур одного типа
Dim shp As Visio.Shape
    
    'Перебираем все фигуры в выделении и если очередная фигура имеет такую же ячейку - присваиваем ей новое значение
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).formula = aB_Value
        End If
    Next shp
    
End Sub

'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub


Public Sub HideMaster(ByVal masterName As String, ByVal visible As Integer)
'Прока скрывает/показыват мастер по имени
Dim mstr As Visio.Master
Dim doc As Visio.Document

    Set doc = ThisDocument
    Set mstr = doc.Masters(masterName)
    mstr.Hidden = Not visible
    
    Set mstr = Nothing
    Set doc = Nothing
End Sub
'HideMaster "Дорога 2", 1 - скрыто
'HideMaster "Дорога 2", 0 - видимо

Public Sub SeekBuilding(ShpObj As Visio.Shape)
'Процедура получения степени огнестойкости и присвоение его фигуре обозначения степени огнестойкости
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim Col As Collection

    On Error GoTo EX
'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'Перебираем все фигуры на странице
    For Each OtherShape In Application.ActivePage.Shapes
        If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.Cells("User.IndexPers") = 135 And OtherShape.HitTest(x, y, 0.01) > 1 Then
                ShpObj.Cells("Prop.SO").FormulaU = _
                 "Sheet." & OtherShape.ID & "!Prop.SO"
            End If
        End If
    Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    SaveLog Err, "SeekFire", ShpObj.Name
End Sub
