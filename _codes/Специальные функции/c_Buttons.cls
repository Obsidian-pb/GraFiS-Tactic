VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Класс для обработки нажатий на кнопки

Public WithEvents ComBut As Office.CommandBarButton   'Экспорт в JPG
Attribute ComBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  'Аспект
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  'Исправить положение
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  'Количество фигур
Attribute ComBut4.VB_VarHelpID = -1
Public WithEvents ComBut5 As Office.CommandBarButton  'Таймер
Attribute ComBut5.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("Спецфункции").Controls("Экспорт в JPG")
    Set ComBut2 = Application.CommandBars("Спецфункции").Controls("Аспект")
    Set ComBut3 = Application.CommandBars("Спецфункции").Controls("Исправить расположение")
    Set ComBut4 = Application.CommandBars("Спецфункции").Controls("Количество фигур")
    Set ComBut5 = Application.CommandBars("Спецфункции").Controls("Таймер")

End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
    Set ComBut2 = Nothing
    Set ComBut3 = Nothing
    Set ComBut4 = Nothing
    Set ComBut5 = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Экспорт в JPG"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    JPGExportAll_P
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Аспект"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    SetAspect_P
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Исправить расположение"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    FixZIndex_P
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Количество фигур"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShapesCountShow
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComBut5_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Таймер"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Показываем/скрываем панель иснтрументов "Таймер"
    If Ctrl.state = msoButtonUp Then
        AddTimer
    Else
        DelTBTimer
    End If
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub


Public Sub PictureRefresh()
'Обновляем картинки на кнопках
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---Обновляем иконки на кнопках
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("Спецфункции").Controls("Экспорт в JPG")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\ExportJPG1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\ExportJPG2.bmp")
    End With
    
    Set cb = Application.CommandBars("Спецфункции").Controls("Аспект")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Aspect1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Aspect2.bmp")
    End With
    
    Set cb = Application.CommandBars("Спецфункции").Controls("Исправить расположение")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Fix1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fix2.bmp")
    End With
    
    Set cb = Application.CommandBars("Спецфункции").Controls("Количество фигур")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Count1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Count2.bmp")
    End With
    
Set cb = Nothing
End Sub
