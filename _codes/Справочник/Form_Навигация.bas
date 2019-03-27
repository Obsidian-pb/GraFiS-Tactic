VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Навигация"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
    s_CheckColor (1)
End Sub

Private Sub ОбластьДанных_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Процедура появления подчеркивания надписей меню при проведении курсора над областью данных
    If П_ГС_ПА.FontUnderline = True Then П_ГС_ПА.FontUnderline = False
    If П_ГС_СПА.FontUnderline = True Then П_ГС_СПА.FontUnderline = False
    If П_ГС_ПрочаяТехника.FontUnderline = True Then П_ГС_ПрочаяТехника.FontUnderline = False
    If П_ГС_Компоненты.FontUnderline = True Then П_ГС_Компоненты.FontUnderline = False
    If П_ГС_ПТВ.FontUnderline = True Then П_ГС_ПТВ.FontUnderline = False
    If П_ГС_ГДЗС.FontUnderline = True Then П_ГС_ГДЗС.FontUnderline = False
    If П_ГС_Водоснабжение.FontUnderline = True Then П_ГС_Водоснабжение.FontUnderline = False
    If П_ГС_Свойства.FontUnderline = True Then П_ГС_Свойства.FontUnderline = False
    If П_ГС_Параметры.FontUnderline = True Then П_ГС_Параметры.FontUnderline = False
    If П_ГС_Гарнизон.FontUnderline = True Then П_ГС_Гарнизон.FontUnderline = False
    
End Sub

'-------------------------Гиперссылки меню----------------------------------------------
'-------------------------Перемещение указателя-----------------------------------------
Private Sub П_ГС_Водоснабжение_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_Водоснабжение.FontUnderline = False Then П_ГС_Водоснабжение.FontUnderline = True
End Sub

Private Sub П_ГС_Гарнизон_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_Гарнизон.FontUnderline = False Then П_ГС_Гарнизон.FontUnderline = True
End Sub

Private Sub П_ГС_ГДЗС_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_ГДЗС.FontUnderline = False Then П_ГС_ГДЗС.FontUnderline = True
End Sub

Private Sub П_ГС_Компоненты_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_Компоненты.FontUnderline = False Then П_ГС_Компоненты.FontUnderline = True
End Sub

Private Sub П_ГС_ПА_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_ПА.FontUnderline = False Then П_ГС_ПА.FontUnderline = True
End Sub

Private Sub П_ГС_Параметры_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_Параметры.FontUnderline = False Then П_ГС_Параметры.FontUnderline = True
End Sub

Private Sub П_ГС_ПрочаяТехника_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_ПрочаяТехника.FontUnderline = False Then П_ГС_ПрочаяТехника.FontUnderline = True
End Sub

Private Sub П_ГС_ПТВ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_ПТВ.FontUnderline = False Then П_ГС_ПТВ.FontUnderline = True
End Sub

Private Sub П_ГС_Свойства_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_Свойства.FontUnderline = False Then П_ГС_Свойства.FontUnderline = True
End Sub

Private Sub П_ГС_СПА_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If П_ГС_СПА.FontUnderline = False Then П_ГС_СПА.FontUnderline = True
End Sub

'-------------------------Щелчок--------------------------------------------------------
Private Sub П_ГС_Водоснабжение_Click()
    В_Водоснабжение.SetFocus
    s_CheckColor (7)
End Sub

Private Sub П_ГС_Гарнизон_Click()
    В_Гарнизон.SetFocus
    s_CheckColor (10)
End Sub

Private Sub П_ГС_ГДЗС_Click()
    В_ГДЗС.SetFocus
    s_CheckColor (6)
End Sub

Private Sub П_ГС_Компоненты_Click()
    В_Компоненты.SetFocus
    s_CheckColor (4)
End Sub

Private Sub П_ГС_ПА_Click()
    В_ПА.SetFocus
    s_CheckColor (1)
End Sub

Private Sub П_ГС_Параметры_Click()
    В_Параметры.SetFocus
    s_CheckColor (9)
End Sub

Private Sub П_ГС_ПрочаяТехника_Click()
    В_ПрочаяТехника.SetFocus
    s_CheckColor (3)
End Sub

Private Sub П_ГС_ПТВ_Click()
    В_ПТВ.SetFocus
    s_CheckColor (5)
End Sub

Private Sub П_ГС_Свойства_Click()
    В_Свойства.SetFocus
    s_CheckColor (8)
End Sub

Private Sub П_ГС_СПА_Click()
    В_ПАСпециальные.SetFocus
    s_CheckColor (2)
End Sub


'-------------------------Гиперссылки во вкладкках---------------------------------------
'-------------------------Правая>>---------------------------------------
Private Sub П_ГС_КПрочейТехнике_Click()
    В_ПрочаяТехника.SetFocus
    s_CheckColor (3)
End Sub

Private Sub П_ГС_КСпециальнымПА_Click()
    В_ПАСпециальные.SetFocus
    s_CheckColor (2)
End Sub

Private Sub П_ГС_ККомпонентам_Click()
    В_Компоненты.SetFocus
    s_CheckColor (4)
End Sub

Private Sub П_ГС_КПТВ_Click()
    В_ПТВ.SetFocus
    s_CheckColor (5)
End Sub

Private Sub П_ГС_КГДЗС_Click()
    В_ГДЗС.SetFocus
    s_CheckColor (6)
End Sub

Private Sub П_ГС_КВодоснабжение_Click()
    В_Водоснабжение.SetFocus
    s_CheckColor (7)
End Sub

Private Sub П_ГС_КСвойства_Click()
    В_Свойства.SetFocus
    s_CheckColor (8)
End Sub

Private Sub П_ГС_КПараметры_Click()
    В_Параметры.SetFocus
    s_CheckColor (9)
End Sub

Private Sub П_ГС_КГарнизон_Click()
    В_Гарнизон.SetFocus
    s_CheckColor (10)
End Sub

'-------------------------<<Левая---------------------------------------
Private Sub П_ГС_КОбщимПА_Click()
    В_ПА.SetFocus
    s_CheckColor (1)
End Sub

Private Sub П_ГС_НСпециальныеПА_Click()
    В_ПАСпециальные.SetFocus
    s_CheckColor (2)
End Sub

Private Sub П_ГС_НПрочая_Click()
    В_ПАСпециальные.SetFocus
    s_CheckColor (3)
End Sub

Private Sub П_ГС_НКомпоненты_Click()
    В_Компоненты.SetFocus
    s_CheckColor (4)
End Sub

Private Sub П_ГС_НПТВ_Click()
    В_ПТВ.SetFocus
    s_CheckColor (5)
End Sub

Private Sub П_ГС_НГДЗС_Click()
    В_ГДЗС.SetFocus
    s_CheckColor (6)
End Sub

Private Sub П_ГС_НВодоснабжение_Click()
    В_Водоснабжение.SetFocus
    s_CheckColor (7)
End Sub

Private Sub П_ГС_НОВ_Click()
    В_Свойства.SetFocus
    s_CheckColor (8)
End Sub

Private Sub П_ГС_НПараметры_Click()
    В_Параметры.SetFocus
    s_CheckColor (9)
End Sub



Private Sub Прямоугольник23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Процедура появления подчеркивания надписей меню при проведении курсора над Прямоугольником23
    If П_ГС_ПА.FontUnderline = True Then П_ГС_ПА.FontUnderline = False
    If П_ГС_СПА.FontUnderline = True Then П_ГС_СПА.FontUnderline = False
    If П_ГС_ПрочаяТехника.FontUnderline = True Then П_ГС_ПрочаяТехника.FontUnderline = False
    If П_ГС_Компоненты.FontUnderline = True Then П_ГС_Компоненты.FontUnderline = False
    If П_ГС_ПТВ.FontUnderline = True Then П_ГС_ПТВ.FontUnderline = False
    If П_ГС_ГДЗС.FontUnderline = True Then П_ГС_ГДЗС.FontUnderline = False
    If П_ГС_Водоснабжение.FontUnderline = True Then П_ГС_Водоснабжение.FontUnderline = False
    If П_ГС_Свойства.FontUnderline = True Then П_ГС_Свойства.FontUnderline = False
    If П_ГС_Параметры.FontUnderline = True Then П_ГС_Параметры.FontUnderline = False
    If П_ГС_Гарнизон.FontUnderline = True Then П_ГС_Гарнизон.FontUnderline = False
    
End Sub

Private Sub s_CheckColor(as_Objectnumber As Integer)
'Процедура выделения активной вкладки на надписях навигационных гиперссылок
    If as_Objectnumber = 1 Then П_ГС_ПА.BackStyle = 1 Else П_ГС_ПА.BackStyle = 0
    If as_Objectnumber = 2 Then П_ГС_СПА.BackStyle = 1 Else П_ГС_СПА.BackStyle = 0
    If as_Objectnumber = 3 Then П_ГС_ПрочаяТехника.BackStyle = 1 Else П_ГС_ПрочаяТехника.BackStyle = 0
    If as_Objectnumber = 4 Then П_ГС_Компоненты.BackStyle = 1 Else П_ГС_Компоненты.BackStyle = 0
    If as_Objectnumber = 5 Then П_ГС_ПТВ.BackStyle = 1 Else П_ГС_ПТВ.BackStyle = 0
    If as_Objectnumber = 6 Then П_ГС_ГДЗС.BackStyle = 1 Else П_ГС_ГДЗС.BackStyle = 0
    If as_Objectnumber = 7 Then П_ГС_Водоснабжение.BackStyle = 1 Else П_ГС_Водоснабжение.BackStyle = 0
    If as_Objectnumber = 8 Then П_ГС_Свойства.BackStyle = 1 Else П_ГС_Свойства.BackStyle = 0
    If as_Objectnumber = 9 Then П_ГС_Параметры.BackStyle = 1 Else П_ГС_Параметры.BackStyle = 0
    If as_Objectnumber = 10 Then П_ГС_Гарнизон.BackStyle = 1 Else П_ГС_Гарнизон.BackStyle = 0
    
    
    
End Sub



