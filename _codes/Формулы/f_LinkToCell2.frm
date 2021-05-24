VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_LinkToCell2 
   Caption         =   "Выбор соединяемых ячеек"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   OleObjectBlob   =   "f_LinkToCell2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_LinkToCell2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private shpTo As Visio.Shape
Private shpFrom As Visio.Shape
Private shpConn As Visio.Shape


Public Sub showForm(ByRef shp1 As Visio.Shape, ByRef shp2 As Visio.Shape, ByRef a_shpConn As Visio.Shape)
Dim i As Integer
    
    'Сохраняем ссылки на фигуры
    Set shpFrom = shp1
    Set shpTo = shp2
    Set shpConn = a_shpConn
    
    'Заполняем список исходных ячеек
    lb_ShapeFrom.Clear
    For i = 0 To shp1.RowCount(visSectionProp) - 1
        lb_ShapeFrom.AddItem shp1.CellsSRC(visSectionProp, i, visTagDefault).RowNameU
        lb_ShapeFrom.List(i, 1) = shp1.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(visUnitsString)
    Next i

    'Заполяем список конечных ячеек
    lb_ShapeTo.Clear
    For i = 0 To shp2.RowCount(visSectionProp) - 1
        lb_ShapeTo.AddItem shp2.CellsSRC(visSectionProp, i, visTagDefault).RowNameU
        lb_ShapeTo.List(i, 1) = shp2.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(visUnitsString)
    Next i
    
    Me.Show
End Sub

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    'Устанваливаем связи
    ConnectShapes
    
    'Закрываем форму
    Me.Hide
End Sub

Private Sub ConnectShapes()
Dim cellFromName As String
Dim cellToName As String
Dim i As Integer
Dim frml As String
    
    'Определяем имя ячейки от которой будет получаться значение
    For i = 0 To Me.lb_ShapeFrom.ListCount - 1
        If Me.lb_ShapeFrom.Selected(i) Then
            cellFromName = Me.lb_ShapeFrom.List(i, 0)
            Exit For
        End If
    Next i
    
    'Определяем имя ячейки которой будет передаваться значение
    For i = 0 To Me.lb_ShapeTo.ListCount - 1
        If Me.lb_ShapeTo.Selected(i) Then
            cellToName = Me.lb_ShapeTo.List(i, 0)
            Exit For
        End If
    Next i
    
'    'Связываем ячейки
'    frml = "Sheet." & shpFrom.ID & "!Prop." & cellFromName
'    '---для случая, если пользователь хочет создать новую ячейку
'    If cellToName = "Новая" Then
'        cellToName = cellToName & "_" & shpFrom.ID
'        shpTo.AddNamedRow visSectionProp, cellToName, visTagDefault
'    End If
'
'    shpTo.Cells("Prop." & cellToName).FormulaU = frml
'
'    'Показываем в коннекторе параметр связи
'    frml = Chr(34) & cellFromName & "=>" & cellToName & ": " & Chr(34) & "&" & frml
'    shpConn.Characters.AddCustomFieldU frml, visFmtNumGenNoUnits
    
    'Связываем ячейки
    frml = "Sheet." & shpFrom.ID & "!Prop." & cellFromName
    '---для случая, если пользователь хочет создать новую ячейку
    If cellToName = "Новая" Then
        cellToName = cellToName & "_" & shpFrom.ID
        shpTo.AddNamedRow visSectionProp, cellToName, visTagDefault
    End If
    
    'Добавляем в коннектор ячейку с переменной
    shpConn.AddNamedRow visSectionProp, cellFromName, visTagDefault
    shpConn.Cells("Prop." & cellFromName).FormulaU = frml
    'Показываем в коннекторе параметр связи
    frml = Chr(34) & cellFromName & "=>" & cellToName & ": " & Chr(34) & "&" & frml
    shpConn.Characters.AddCustomFieldU frml, visFmtNumGenNoUnits
    
    'Добавляем конечной фигуре ссылку на ячейку с переменной в коннекторе
    frml = "Sheet." & shpConn.ID & "!Prop." & cellFromName
    shpTo.Cells("Prop." & cellToName).FormulaU = frml
    
    
End Sub
