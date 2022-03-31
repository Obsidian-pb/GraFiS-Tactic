Attribute VB_Name = "m_ExportDescriptions"
Option Explicit

'Private exl As Object ' Excel.Application
'Private wkbk As Object '  Excel.Workbook
'Private wkst As Object '  Excel.Worksheet
Private cmndID As Integer




Public Sub DescriptionExportToWord()
'Экспортируем описание боевых действий в документ Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
Dim wrdTblRow As Object
'Dim comRowsCount As Integer
Dim i As Integer
Dim r As Integer
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
'    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, comColSorted.Count, 9)
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, 1, 9)
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
        r = 1
        A.Refresh Application.ActivePage.Index, curTime
'        fireTime = A.Result("FireTime")     'Не слишкои надежно
        fireTime = cellval(Application.ActivePage.Shapes, "Prop.FireTime", visUnitsString)  ' visDate)    'Так надежнее
        curTime = comColSorted(1).time
        '---Вставка первой записи
        If DateDiff("n", fireTime, curTime) < 2000 Then
            wrdTbl.Rows(r).Cells(1).Range.text = "Ч+" & DateDiff("n", fireTime, curTime)
        Else
            wrdTbl.Rows(r).Cells(1).Range.text = Format(curTime, "HH:MM")
        End If
        wrdTbl.Rows(r).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
        wrdTbl.Rows(r).Cells(4).Range.text = A.Result("StvolWBHave")
        wrdTbl.Rows(r).Cells(5).Range.text = A.Result("StvolWAHave")
        wrdTbl.Rows(r).Cells(6).Range.text = A.Result("StvolWLHave")
        wrdTbl.Rows(r).Cells(7).Range.text = A.Result("StvolFoamHave")
        wrdTbl.Rows(r).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
        
        
        For i = 1 To comColSorted.Count
            Set com = comColSorted(i)
            'Если время текущей фигуры больше чем время предыдущей - обновляем анализатор
            If curTime <> com.time Then
                A.Refresh Application.ActivePage.Index, com.time
                '---Вставка записей начиная со второй (при условии их изменения)
                r = wrdTbl.Rows.Add().Index
                
                wrdTbl.Rows(r).Cells(1).Range.text = "Ч+" & DateDiff("n", fireTime, com.time)
                wrdTbl.Rows(r).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
                wrdTbl.Rows(r).Cells(4).Range.text = A.Result("StvolWBHave")
                wrdTbl.Rows(r).Cells(5).Range.text = A.Result("StvolWAHave")
                wrdTbl.Rows(r).Cells(6).Range.text = A.Result("StvolWLHave")
                wrdTbl.Rows(r).Cells(7).Range.text = A.Result("StvolFoamHave")
                wrdTbl.Rows(r).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
                
                curTime = comColSorted(i).time
            End If
            
            If com.sdType = 1 Then
                If Asc(wrdTbl.Rows(r).Cells(9).Range.text) = 13 Then
                    wrdTbl.Rows(r).Cells(9).Range.text = com.text
                Else
                    wrdTbl.Rows(r).Cells(9).Range.text = wrdTbl.Rows(r).Cells(9).Range.text & Chr(10) & com.text
                End If
            Else
                If Asc(wrdTbl.Rows(r).Cells(9).Range.text) = 13 Then
                    wrdTbl.Rows(r).Cells(2).Range.text = com.text
                Else
                    wrdTbl.Rows(r).Cells(2).Range.text = wrdTbl.Rows(r).Cells(2).Range.text & Chr(10) & com.text
                End If
            End If
            
        Next i
    End If
    wrdTbl.AutoFitBehavior 1        'Устанавливаем ширину столбцов по содержимому
End Sub

Public Sub DescriptionViewInList()
'Экспортируем описанме боевых действий в документ Word
Dim i As Integer
Dim shp As Visio.Shape
Dim comCol As Collection
Dim comColSorted As Collection
Dim com As c_SimpleDescription
Dim curTime As Date
Dim fireTime As Date

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set comCol = New Collection
    cmndID = 0
    For Each shp In Application.ActivePage.Shapes
        checkCommands comCol, shp
    Next shp
    
    
    
    '---Сортируем
    Set comColSorted = SortCommands(comCol)
    
  
    
    'Заполняем таблицу описания боевых действий
    If comColSorted.Count > 0 Then
        A.Refresh Application.ActivePage.Index, curTime
        
        fireTime = cellval(Application.ActivePage.Shapes, "Prop.FireTime", visUnitsString)  ' visDate)    'Так надежнее
'        fireTime = A.Result("FireTime")
        curTime = comColSorted(1).time
        
        ReDim myArray(comColSorted.Count, 3)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Время"
        myArray(0, 2) = "Изменение обстановки"
        myArray(0, 3) = "Команда/Действие"
    
        For i = 1 To comColSorted.Count
            Set com = comColSorted(i)
                curTime = comColSorted(i).time
                
                myArray(i, 0) = com.shp.ID
                If DateDiff("n", fireTime, curTime) < 2000 Then
                    myArray(i, 1) = "Ч+" & DateDiff("n", fireTime, curTime)
                Else
                    myArray(i, 1) = Format(curTime, "HH:MM")
                End If
                
                If com.sdType = 1 Then
                    myArray(i, 3) = com.text
                Else
                    myArray(i, 2) = com.text
                End If
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;300 pt;300 pt", "InfoAndCommands", "Обстановка и команды"

End Sub










Private Sub checkCommands(ByRef comCol As Collection, ByRef shp As Visio.Shape)
Dim i As Integer
Dim rowName As String
Dim cmnd As c_SimpleDescription
    
    On Error GoTo ex
    
    For i = 0 To shp.RowCount(visSectionUser) - 1
        rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
        If Len(rowName) > 9 Then
            If left(rowName, 12) = "GFS_Command_" Then
                Set cmnd = New c_SimpleDescription
                cmnd.Activate shp, shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), CStr(cmndID)
                AddUniqueCollectionItem comCol, cmnd
'                        shp.Cells("Actions." & rowName & ".Action").Formula = Replace(shp.Cells("Actions." & rowName & ".Action").Formula, "РТП", "Управление_СиС")
                
                cmndID = cmndID + 1
            End If
            
            If left(rowName, 9) = "GFS_Info_" Then
                Set cmnd = New c_SimpleDescription
                cmnd.ActivateAsInfo shp, shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), CStr(cmndID)
                AddUniqueCollectionItem comCol, cmnd
'                        shp.Cells("Actions." & rowName & ".Action").Formula = Replace(shp.Cells("Actions." & rowName & ".Action").Formula, "РТП", "Управление_СиС")

                
                cmndID = cmndID + 1
            End If
        End If
    Next i
    
    
Exit Sub
ex:

End Sub



Public Function getCallName(ByRef shp As Visio.Shape) As String
Dim txt As String

    On Error GoTo ex
    getCallName = ""
    'Подразделение
    txt = shp.Cells("Prop.Unit").ResultStr(visUnitsString)
    If Not txt = "" Then
        getCallName = txt
    End If
    'Позывной
    txt = shp.Cells("Prop.Call").ResultStr(visUnitsString)
    If Not txt = "" Then
        getCallName = getCallName & "(" & txt & ")"
    End If
    
Exit Function
ex:
    getCallName = "-"
End Function


