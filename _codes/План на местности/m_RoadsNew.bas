Attribute VB_Name = "m_RoadsNew"
'Новый модуль для функций по работе с новыми фигурами дорог
Option Explicit




Sub PS_FormRoad(ShpObj As Visio.Shape)
'Прока добавляет к лицевой части фоновую и привязывает ее к лицевой
Dim FaceRoad As Visio.Shape        'Лицевая проезжая часть
Dim BckRoad As Visio.Shape         'подложка
Dim MasterBckRoad As Visio.Shape   'мастер подложки
Dim str As String

On Error GoTo EX
'1 запоминаем фигуры
    Set MasterBckRoad = ThisDocument.Masters("Bckgnd").Shapes(1)
    Set FaceRoad = ShpObj
    Set BckRoad = Application.ActivePage.Drop(MasterBckRoad, 0, 0)

'2 Присваиваем новой фигуре требуемые для нее свойства
    BckRoad.Cells("LineColor").FormulaForce = "GUARD(0)"
'    BckRoad.Cells("EventXFMod").FormulaForce = "CALLTHIS(" & Chr(34) & "PS_SendMeBack" & Chr(34) & ";" & Chr(34) & "План_на_местности" & Chr(34) & ")"
    BckRoad.Cells("Scratch.A1").Formula = "ISERR(" & FaceRoad.NameID & "!PinX)"
    BckRoad.Cells("Scratch.B1").Formula = "DEPENDSON(Scratch.A1)+CALLTHIS(" & Chr(34) & "PS_CheckDeletion" & Chr(34) & ";" & Chr(34) & "План_на_местности" & Chr(34) & ")"
    
'3 связываем фигуру фона с лицевой
    BckRoad.Cells("BeginX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!BeginX)"
    BckRoad.Cells("BeginY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!BeginY)"
    BckRoad.Cells("EndX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!EndX)"
    BckRoad.Cells("EndY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!EndY)"
    BckRoad.Cells("LineWeight").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LineWeight+1pt)"
    BckRoad.Cells("LineCap").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LineCap+1pt)"
    
'4 возвращаем фокус лицевой фигуре
    ActiveWindow.DeselectAll
    ActiveWindow.Select FaceRoad, visSelect

EX:
Exit Sub
    SaveLog Err, "PS_FormRoad"
End Sub

Sub PS_FormRoad2(ShpObj As Visio.Shape)
'Прока добавляет к лицевой части фоновую и привязывает ее к лицевой
'Прока аналогична PS_FormRoad, но работает с фигурами построенными на Polyline
Dim FaceRoad As Visio.Shape        'Лицевая проезжая часть
Dim BckRoad As Visio.Shape         'подложка
Dim MasterBckRoad As Visio.Shape   'мастер подложки
Dim str As String

    On Error GoTo EX
'1 запоминаем фигуры
    Set MasterBckRoad = ThisDocument.Masters("Bckgnd2").Shapes(1)
    Set FaceRoad = ShpObj
    Set BckRoad = Application.ActivePage.Drop(MasterBckRoad, 0, 0)

'2 Присваиваем новой фигуре требуемые для нее свойства
    BckRoad.Cells("LineColor").FormulaForce = "GUARD(0)"
    BckRoad.Cells("EventXFMod").FormulaForce = "CALLTHIS(" & Chr(34) & "PS_SendMeBack" & Chr(34) & ";" & Chr(34) & "План_на_местности" & Chr(34) & ")"
    BckRoad.Cells("Scratch.A1").Formula = "ISERR(" & FaceRoad.NameID & "!PinX)"
    BckRoad.Cells("Scratch.B1").Formula = "DEPENDSON(Scratch.A1)+CALLTHIS(" & Chr(34) & "PS_CheckDeletion" & Chr(34) & ";" & Chr(34) & "План_на_местности" & Chr(34) & ")"
    
'3 связываем фигуру фона с лицевой
    BckRoad.Cells("Width").FormulaForce = "GUARD(" & FaceRoad.NameID & "!Width)"
    BckRoad.Cells("Height").FormulaForce = "GUARD(" & FaceRoad.NameID & "!Height)"
    BckRoad.Cells("Angle").FormulaForce = "GUARD(" & FaceRoad.NameID & "!Angle)"
    BckRoad.Cells("PinX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!PinX)"
    BckRoad.Cells("PinY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!PinY)"
    BckRoad.Cells("LocPinX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LocPinX)"
    BckRoad.Cells("LocPinY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LocPinY)"
'    BckRoad.Cells("BeginX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!BeginX)"
'    BckRoad.Cells("BeginY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!BeginY)"
'    BckRoad.Cells("EndX").FormulaForce = "GUARD(" & FaceRoad.NameID & "!EndX)"
'    BckRoad.Cells("EndY").FormulaForce = "GUARD(" & FaceRoad.NameID & "!EndY)"
    BckRoad.Cells("LineWeight").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LineWeight+1pt)"
    BckRoad.Cells("LineCap").FormulaForce = "GUARD(" & FaceRoad.NameID & "!LineCap+1pt)"
    'Связываем геометрию
    BckRoad.CellsSRC(visSectionFirstComponent, 1, 0).FormulaForce = "GUARD(" & FaceRoad.NameID & "!Geometry1.X1)"
    BckRoad.CellsSRC(visSectionFirstComponent, 1, 1).FormulaForce = "GUARD(" & FaceRoad.NameID & "!Geometry1.Y1)"
    BckRoad.CellsSRC(visSectionFirstComponent, 2, 0).FormulaForce = "GUARD(" & FaceRoad.NameID & "!Geometry1.X2)"
    BckRoad.CellsSRC(visSectionFirstComponent, 2, 1).FormulaForce = "GUARD(" & FaceRoad.NameID & "!Geometry1.Y2)"
    BckRoad.CellsSRC(visSectionFirstComponent, 2, 2).FormulaForce = "GUARD(" & FaceRoad.NameID & "!Geometry1.A2)"
    
    'FaceRoad.CellsSRC(visSectionFirstComponent, 2, 2).FormulaForce = FaceRoad.CellsSRC(visSectionFirstComponent, 2, 2).FormulaU
    
'4 возвращаем фокус лицевой фигуре
    ActiveWindow.DeselectAll
    ActiveWindow.Select FaceRoad, visSelect

Exit Sub
EX:
    SaveLog Err, "PS_FormRoad2"
End Sub


Sub PS_CheckDeletion(ShpObj As Visio.Shape)
'Прока проверяет, не была ли удалена фигура лицевой стороны, и если была, то удаляется сама
    On Error GoTo EX
    
    If ShpObj.Cells("Scratch.A1").Result(visNone) = 1 Then
        ShpObj.Delete
    End If
Exit Sub
EX:
    
End Sub

Sub PS_SendMeBack(ShpObj As Visio.Shape)
'Отправляет текущую фигуру назад
    ShpObj.SendToBack
End Sub
