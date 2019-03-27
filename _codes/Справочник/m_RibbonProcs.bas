Attribute VB_Name = "m_RibbonProcs"
Option Compare Database

Public Sub GetImageSplash(control As IRibbonControl, ByRef image)
'Процедура присваивает иконку для кнопки "Показать форму заставки"
Dim pth As String

pth = Application.CurrentProject.path & "\Bitmaps\Zastavka.jpg"
    Set image = LoadPicture(pth)

End Sub

Public Sub GetImageNav(control As IRibbonControl, ByRef image)
'Процедура присваивает иконку для кнопки "Показать форму навигации"
Dim pth As String

pth = Application.CurrentProject.path & "\Bitmaps\Nav.jpg"
    Set image = LoadPicture(pth)

End Sub

Public Sub GetImageImport(control As IRibbonControl, ByRef image)
'Процедура присваивает иконку для кнопки "Импортировать данные"
Dim pth As String

pth = Application.CurrentProject.path & "\Bitmaps\DataImport.jpg"
    Set image = LoadPicture(pth)

End Sub

Public Sub GetImageHelp(control As IRibbonControl, ByRef image)
'Процедура присваивает иконку для кнопки "Показать справку"
Dim pth As String

pth = Application.CurrentProject.path & "\Bitmaps\Spravka.jpg"
    Set image = LoadPicture(pth)

End Sub

'--------------------Копирование записей в дргие наборы-----------------------------
Public Sub GetImageCopytoSet(control As IRibbonControl, ByRef image)
'Прока присваивает иконку для кнопки "Скопировать в другой набор"
Dim pth As String

pth = Application.CurrentProject.path & "\Bitmaps\CopyToSet.jpg"
    Set image = LoadPicture(pth)
End Sub
