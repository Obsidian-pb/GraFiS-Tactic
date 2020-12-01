Attribute VB_Name = "m_WorkWithHTML"
Option Explicit



Public Sub ShowData(ByRef shp As Visio.Shape, ByVal htmlText As String)
    f_FormulaForm.ShowHTML shp, htmlText
End Sub

Public Sub ShowDataInShape(ByRef shp As Visio.Shape, ByVal htmlText As String)
'    shp.Text = ""
    f_FormulaForm.CopyBrowserContent shp, htmlText
    SetShowDataInShapeControlFormula shp, "User.DataChangeAction.Prompt"
End Sub

Private Sub SetShowDataInShapeControlFormula(ByRef shp As Visio.Shape, ByVal cellName As String)
Dim i As Integer
Dim frml As String
Dim nameOfRow As String
    
    frml = " "
    For i = 0 To shp.RowCount(visSectionProp) - 1
        'Получаем имя строки
        nameOfRow = shp.CellsSRC(visSectionProp, i, visCustPropsValue).RowNameU
               
        'Составляем итоговую формулу для контроля изменений (для отраджения в тексте фигуры)
        If Len(frml) > 0 Then frml = frml & "&"
        frml = frml & "Prop." & nameOfRow
    Next i
    
    SetCellFrml shp, cellName, frml
End Sub

Public Sub ClearShapeText(ByRef shp As Visio.Shape)
    shp.Text = " "
End Sub

Public Sub TryGetFromAnalaizer(ByRef shp As Visio.Shape, ByVal nameOfRow As String)
    If left(nameOfRow, 2) = "A_" Then
        SetCellVal shp, "Prop." & nameOfRow, a.Result(Right(nameOfRow, Len(nameOfRow) - 2))
    End If
End Sub


Public Sub TTT()
Dim arr  As Variant
'Использовать это для получения списка элементов анализа в форме добавления ячеек для анализа
    arr = a.GetElementsCode
    
    Debug.Print arr(1, 0), arr(1, 1)
End Sub


