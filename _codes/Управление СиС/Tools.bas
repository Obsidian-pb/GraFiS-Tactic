Attribute VB_Name = "Tools"
Option Explicit

Public Function GetScaleAt200() As Double
'Возвращает коэффициент приведения размера текущей страницы относительно масштаба 1:200
Dim v_Minor As Double
Dim v_Major As Double

    v_Minor = Application.ActivePage.PageSheet.Cells("PageScale").Result(visNumber)
    v_Major = Application.ActivePage.PageSheet.Cells("DrawingScale").Result(visNumber)
    GetScaleAt200 = (v_Major / v_Minor) / 200
End Function

Public Function GetGFSShapeTime(ByRef shp As Visio.Shape) As Double
    
    GetGFSShapeTime = cellVal(shp, "Prop.SetTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellVal(shp, "Prop.FormingTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellVal(shp, "Prop.ArrivalTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function

GetGFSShapeTime = 0
End Function

Public Sub P_TryDeleteSmartTag(ByRef shp As Visio.Shape, stName As String, rowName As String)
'stName - название смарт-тега, rowName - название строки смарт-тега в секции SmartTags
Dim i As Integer
Dim smartTagRowIndex As Integer
    
    'Получаем инекс смарт тега и проверяем имеется ли такой смарт тег
    smartTagRowIndex = P_GetRowIndex(shp, rowName)
    If smartTagRowIndex >= 0 Then
        'Проверяем все строки секции Actions на предмет наличия ссылок на указанный смарт тег
        For i = 0 To shp.RowCount(visSectionAction) - 1
            'Если есть хоть одна - выходим из процедуры не удаляя смарт тег
            If shp.CellsSRC(visSectionAction, i, visActionTagName).ResultStr(visUnitsString) = stName Then Exit Sub
        Next i
    
        shp.DeleteRow visSectionSmartTag, smartTagRowIndex
        'Удаляем так же и ячейку отслеживания времени
        On Error Resume Next
        shp.DeleteRow visSectionUser, shp.Cells("User.CurrentDocTime").row
    End If

End Sub

Public Function P_GetRowIndex(ByRef shp As Visio.Shape, cellName As String) As Integer
    On Error GoTo ex
    P_GetRowIndex = shp.Cells(cellName).row
Exit Function
ex:
    P_GetRowIndex = -1
End Function

'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub



