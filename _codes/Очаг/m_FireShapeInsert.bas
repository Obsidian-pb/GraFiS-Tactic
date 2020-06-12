Attribute VB_Name = "m_FireShapeInsert"
Option Explicit
'--------------------------------------Модуль добавленияплощадей горения по указанным данным----------------------


Public Sub Sm_ShapeFormShow(ShpObj As Visio.Shape)
'Процедура показвает форму добавления площадей горения в соответствии с заданными показателями

    On Error GoTo EX
'---Определяем стартовые значения формы
'    F_InsertFire.TB_Time.Value = ShpObj.Cells("Prop.FireTime").ResultStr(visDate)
    F_InsertFire.TB_Time.value = ActiveDocument.DocumentSheet.Cells("User.CurrentTime").ResultStr(0)
    F_InsertFire.TB_Duration.value = DateDiff("n", ActiveDocument.DocumentSheet.Cells("User.FireTime").Result(visDate), _
                                        ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate))
    F_InsertFire.TB_Radius.value = Round(ShpObj.Shapes.item(4).Cells("Width").Result(visMeters), 2)

'---Указываем объекту формы, с каким ID объект  его вызвал
    F_InsertFire.Vfl_TargetShapeID = ShpObj.ID

'---Указываем объекту формы, стартовую дату
    F_InsertFire.VmD_TimeStart = ActiveDocument.DocumentSheet.Cells("User.FireTime").Result(visDate)
    F_InsertFire.FireTime.Caption = "Начало пожара: " & ActiveDocument.DocumentSheet.Cells("User.FireTime").ResultStr(0)

'---Показываем форму
    F_InsertFire.Show

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "Sm_ShapeFormShow"
End Sub

Public Sub Sm_ExtSquareFormShow(ShpObj As Visio.Shape)
'Процедура показвает форму расчета площади тушения

    On Error GoTo EX

'---Указываем объекту формы, какой объект его вызвал
    F_InsertExtSquare.SetFireShape ShpObj

'---Показываем форму
    F_InsertExtSquare.Show

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "Sm_ExtSquareFormShow"
End Sub

