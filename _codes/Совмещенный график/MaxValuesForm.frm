VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MaxValuesForm 
   Caption         =   "Укажите новые максимальные значения"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3510
   OleObjectBlob   =   "MaxValuesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MaxValuesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetShape As Visio.Shape


Public Sub PS_ShowME(ByRef shp As Visio.Shape)
'Прока показывающая форму и устаналвивающая значения по умолчанию
    
    
'---Задаем значения по умолчанию
    Me.TB_MaxSquare.Text = shp.Cells("Prop.FireMax") '.ResultStr(visUnitsString)
    Me.TB_MaxTime.Text = shp.Cells("Prop.TimeMax") '.ResultStr(visUnitsString)
    
'---Сохраняем целевую фигуру
    Set TargetShape = shp
    
'---Показываем форму и
    Me.Show

End Sub

Private Sub CB_Cancel_Click()
    Set TargetShape = Nothing
    Me.Hide
End Sub

Private Sub CB_OK_Click()
    '---Обновляем данные
    TargetShape.Cells("Prop.FireMax").FormulaForceU = "GUARD(" & Me.TB_MaxSquare.Text & ")"
    TargetShape.Cells("Prop.TimeMax").FormulaForceU = "GUARD(" & Me.TB_MaxTime.Text & ")"
    
    '---Закрываем форму
    Set TargetShape = Nothing
    Me.Hide
End Sub
