Attribute VB_Name = "m_ExportDescriptions"
Option Explicit

Private exl As Object ' Excel.Application
Private wkbk As Object '  Excel.Workbook
Private wkst As Object '  Excel.Worksheet
Private cmndID As Integer




Public Sub DescriptionExportToWord()
'Экспортируем описанме боевых действий в документ Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
'Dim comRowsCount As Integer
Dim i As Integer
Dim shp As Visio.Shape
Dim comCol As Collection
Dim comColSorted As Collection
Dim com As c_SimpleDescription
Dim curTime As Date
Dim fireTime As Date
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set comCol = New Collection
    cmndID = 0
    For Each shp In Application.ActivePage.Shapes
        checkCommands comCol, shp
    Next shp
    
    '---Сортируем
    Set comColSorted = SortCommands(comCol)
    
    
    'Создаем новый документ Word
    Set wrd = CreateObject("Word.Application")
    wrd.visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add
    wrdDoc.Activate
    'Создаем в новом документе таблицу требуемых размеров
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, comColSorted.Count, 9)
    With wrdTbl
        If .style <> "Сетка таблицы" Then
            .style = "Сетка таблицы"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    
    
    'Заполняем таблицу описания боевых действий
    If comColSorted.Count > 0 Then
        A.Refresh Application.ActivePage.Index, curTime
'        fireTime = A.Result("FireTime")     'Не слишкои надежно
        fireTime = cellVal(Application.ActivePage.Shapes, "Prop.FireTime", visDate)    'Так надежнее
        curTime = comColSorted(1).time
        '---Вставка первой записи
        wrdTbl.Rows(1).Cells(1).Range.text = "Ч+" & DateDiff("n", fireTime, curTime)
        wrdTbl.Rows(1).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
        wrdTbl.Rows(1).Cells(4).Range.text = A.Result("StvolWBHave")
        wrdTbl.Rows(1).Cells(5).Range.text = A.Result("StvolWAHave")
        wrdTbl.Rows(1).Cells(6).Range.text = A.Result("StvolWLHave")
        wrdTbl.Rows(1).Cells(7).Range.text = A.Result("StvolFoamHave")
        wrdTbl.Rows(1).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
                
        
        
        For i = 1 To comColSorted.Count
            Set com = comColSorted(i)
            'Если время текущей фигуры больше чем время предыдущей - обновляем канализатор
            If curTime <> com.time Then
                A.Refresh Application.ActivePage.Index, com.time
                '---Вставка записей начиная со второй (при условии их изменения)
                wrdTbl.Rows(i).Cells(1).Range.text = "Ч+" & DateDiff("n", fireTime, com.time)
                wrdTbl.Rows(i).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
                wrdTbl.Rows(i).Cells(4).Range.text = A.Result("StvolWBHave")
                wrdTbl.Rows(i).Cells(5).Range.text = A.Result("StvolWAHave")
                wrdTbl.Rows(i).Cells(6).Range.text = A.Result("StvolWLHave")
                wrdTbl.Rows(i).Cells(7).Range.text = A.Result("StvolFoamHave")
                wrdTbl.Rows(i).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
                
                curTime = comColSorted(i).time
            End If

            wrdTbl.Rows(i).Cells(9).Range.text = com.text
        Next i
    End If
    wrdTbl.AutoFitBehavior 1        'Устанавливаем ширину столбцов по содержимому
End Sub




'Public Sub Export()
'
'
'Dim shp As Visio.Shape
'
'
'    Set exl = CreateObject("Excel.Application")   ' New Excel.Application
'    Set wkbk = exl.Workbooks.Add()
'    Set wkst = exl.ActiveSheet
'    exl.visible = True
'
'    rowNumber = 1
'    For Each shp In Application.ActivePage.Shapes
'        fillCommand shp
'    Next shp
'
''    rowNumber = 1
''    For Each shp In Application.ActivePage.Shapes
''        getSetTime shp
''
'''        rowNumber = rowNumber + 1
''    Next shp
'
'End Sub

Private Sub checkCommands(ByRef comCol As Collection, ByRef shp As Visio.Shape)
Dim i As Integer
Dim rowName As String
Dim cmnd As c_SimpleDescription
    
    On Error GoTo ex
    
    For i = 0 To shp.RowCount(visSectionUser) - 1
        rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
        If Len(rowName) > 12 Then
            If left(rowName, 12) = "GFS_Command_" Then
                Set cmnd = New c_SimpleDescription
                cmnd.Activate shp, shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), CStr(cmndID)
                AddUniqueCollectionItem comCol, cmnd
                
                cmndID = cmndID + 1
            End If
        End If
    Next i
    
    
Exit Sub
ex:

End Sub

'Private Sub fillCommand(ByRef shp As Visio.Shape)
'Dim i As Integer
'Dim rowName As String
'Dim comTime As String
'Dim comText As String
'Dim comArr() As String
'
'    On Error GoTo ex
'
'    For i = 0 To shp.RowCount(visSectionUser) - 1
'        rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
'        If Len(rowName) > 12 Then
'            If left(rowName, 12) = "GFS_Command_" Then
'                comArr = Split(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), delimiter)
'                wkst.Cells(rowNumber, 1) = comArr(0)
'                wkst.Cells(rowNumber, 2) = getCallName(shp) & " " & comArr(UBound(comArr))
'
'                rowNumber = rowNumber + 1
'            End If
'        End If
'
'
'    Next i
'
'
'Exit Sub
'ex:
'
'End Sub

Public Function getCallName(ByRef shp As Visio.Shape) As String
    On Error GoTo ex
    getCallName = shp.Cells("Prop.Call").ResultStr(visUnitsString)
Exit Function
ex:
    getCallName = "-"
End Function
'Public Sub getSetTime(ByRef shp As Visio.Shape)
'    On Error GoTo ex
'    If shp.Cells("User.IndexPers").Result(visNumber) = 34 Then
'        wkst.Cells(rowNumber, 4) = shp.Cells("Prop.SetTime").ResultStr(visDate)
'        wkst.Cells(rowNumber, 5) = shp.Cells("User.DiameterIn").ResultStr(visUnitsString)
'        rowNumber = rowNumber + 1
'    End If
'    If shp.Cells("User.IndexPers").Result(visNumber) = 36 Then
'        wkst.Cells(rowNumber, 4) = shp.Cells("Prop.SetTime").ResultStr(visDate)
'        wkst.Cells(rowNumber, 5) = shp.Cells("User.DiameterIn").ResultStr(visUnitsString)
'        rowNumber = rowNumber + 1
'    End If
'
'
'Exit Sub
'ex:
'
'End Sub

