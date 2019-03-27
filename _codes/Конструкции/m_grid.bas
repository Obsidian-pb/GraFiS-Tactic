Attribute VB_Name = "m_grid"
Option Explicit

'------------------------Модуль для хранения процедур отрисовки решетки плана-------------------

Public Sub DropGridMain(ShpObj As Visio.Shape)
'Процедура построения координатной сетки
On Error GoTo EX

    InsertGridForm.TB_X.Text = ShpObj.Cells("PinX").Result(visMillimeters)
    InsertGridForm.TB_Y.Text = ShpObj.Cells("PinY").Result(visMillimeters)
    InsertGridForm.Show

EX:
    ShpObj.Delete
End Sub

