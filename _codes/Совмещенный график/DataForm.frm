VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataForm 
   Caption         =   "Таблица данных"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   OleObjectBlob   =   "DataForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RefreshNeed As Boolean   'Опция показывающая, нужно ли обновлять данные графика или нет
Public DataCorrect As Boolean   'Верно ли внесены данные
Private TargetShape As Visio.Shape  'Целевая фигура

Const TextFieldWidth As Integer = 42  'Ширина текстового поля


Dim Time As String
Dim SqrExp As String
Dim TimeA() As String
Dim SqrExpA() As String

Dim c_TableColumn() As c_TableColumn


Public Sub ShowMe(ByRef shp As Visio.Shape)
'Основная прока показа формы и установления значений таблицы
Dim i As Integer
Dim IndexPers As Integer

    On Error GoTo EX

    '---Получаем данные фигуры (времена и второе значение)
    SqrExp = shp.Cells("Scratch.A1").ResultStr(visUnitsString)
    Time = shp.Cells("Scratch.B1").ResultStr(visUnitsString)
    StringToArray SqrExp, ";", SqrExpA()
    StringToArray Time, ";", TimeA()
    
    '---Определяем для какого графика осуществляются вычисления и подписываем
    If shp.Cells("User.IndexPers") = 123 Or shp.Cells("User.IndexPers") = 124 Then
        '---Для площадей
        Me.Label2 = "Площадь м.кв."
    ElseIf shp.Cells("User.IndexPers") = 125 Or shp.Cells("User.IndexPers") = 126 Then
        '---Для расходов
        Me.Label2 = "Расход л/с"
    End If
    '---Переопределяем массивы объектных переменных
        ReDim c_TableColumn(UBound(TimeA))
    
    '---Формируем таблицу
        ps_FillTable
    
    
    '---Показываем подготовленную форму
    Me.Show
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ShowMe"
End Sub




'------------------------------Процедуры изменения таблицы----------------------------------------------------
Private Sub ps_FillTable()
'Прока заполнения таблицы
Dim i As Integer

    On Error GoTo EX

    '---В соответсвии с количеством узлов создаем необходимое количество тектовых полей
    For i = 0 To UBound(TimeA)
        Set c_TableColumn(i) = New c_TableColumn
        c_TableColumn(i).Activate i, 60 + TextFieldWidth * i, 6, TimeA(i), SqrExpA(i)
        
        '---Перемещаем кнопку "Добавить" в конец
        Me.CB_Add.Left = 60 + TextFieldWidth * (i + 1)

        '---Устанавливаем размер формы по количеству элементов
        If i < 5 Then
            Me.Width = TextFieldWidth * (5) + 36 + 54
        Else
            Me.Width = TextFieldWidth * (i + 1) + 36 + 54
        End If

        '---Располагаем кнопки посередине формы
        CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
        CB_Cancel.Left = (Me.Width / 2) + 3
    Next i

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ps_FillTable"
End Sub

Private Sub ps_ClearTable()
'Прока очистки таблицы
Dim i As Integer

    On Error GoTo EX

    '---В соответсвии с количеством узлов создаем необходимое количество тектовых полей
    For i = 0 To UBound(TimeA)
        Set c_TableColumn(i) = Nothing
        
        '---Перемещаем кнопку "Добавить" в начало
        Me.CB_Add.Left = TextFieldWidth * (i + 1) + 6

        '---Устанавливаем размер формы по количеству элементов
        If i < 5 Then
            Me.Width = TextFieldWidth * (5) + 36
        Else
            Me.Width = TextFieldWidth * (i + 1) + 36
        End If

        '---Располагаем кнопки посередине формы
        CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
        CB_Cancel.Left = (Me.Width / 2) + 3
    Next i

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ps_ClearTable"
End Sub

Public Sub DeleteColumn(ByVal a_index As Integer)
'Прока удаления указанного столбца
Dim i As Integer

    On Error GoTo EX

    '---Изменяем массив данных для Времени/Значения
        For i = a_index To UBound(TimeA) - 1
            TimeA(i) = TimeA(i + 1)
            SqrExpA(i) = SqrExpA(i + 1)
        Next i
        ReDim Preserve TimeA(UBound(TimeA) - 1)
        ReDim Preserve SqrExpA(UBound(SqrExpA) - 1)
        
    '---Очищаем таблицу
        ps_ClearTable
    
    '---Изменяем массив контролв класса Колонок таблицы
        ReDim c_TableColumn(UBound(TimeA))
        
    '---Заполняем таблицу по новой
        ps_FillTable
        
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "DeleteColumn"
End Sub

Public Sub PS_ChangeTimeValue(ByVal a_ind As Integer, ByVal a_value As String)
    TimeA(a_ind) = a_value
End Sub

Public Sub PS_ChangeDataValue(ByVal a_ind As Integer, ByVal a_value As String)
    SqrExpA(a_ind) = a_value
End Sub

Public Sub PS_GetMainArray(ByRef MainArray())
'Главная прока Формирует рабочий массив
Dim i As Integer
Dim NodeCount As Integer
    
    '---Определяем количество элементов в новом массиве
        NodeCount = UBound(TimeA)
    
    '---Определяем размер нового массива
    ReDim MainArray(1, NodeCount)
    '---Изменяем массив данных для Времени/Значения
    For i = 0 To NodeCount
        '---Время
        MainArray(0, i) = TimeA(i) * 60
        '---Данные
        MainArray(1, i) = SqrExpA(i)
    Next i

End Sub

Private Sub CB_Add_Click()
'Добавляем новую точку
Dim lastitem As Integer

    On Error GoTo EX

    lastitem = UBound(c_TableColumn) + 1

    '---Увеличиваем размеры массивов
    ReDim Preserve TimeA(lastitem)
    ReDim Preserve SqrExpA(lastitem)
    ReDim Preserve c_TableColumn(lastitem)
    
    TimeA(lastitem) = 0
    SqrExpA(lastitem) = 0

    Set c_TableColumn(lastitem) = New c_TableColumn
    c_TableColumn(lastitem).Activate lastitem, 60 + TextFieldWidth * lastitem, 6, TimeA(lastitem), SqrExpA(lastitem)
    
    '---Перемещаем кнопку "Добавить" в конец
    Me.CB_Add.Left = 60 + TextFieldWidth * (lastitem + 1)

    '---Устанавливаем размер формы по количеству элементов
    If lastitem < 5 Then
        Me.Width = TextFieldWidth * (5) + 36 + 54
    Else
        Me.Width = TextFieldWidth * (lastitem + 1) + 36 + 54
    End If

    '---Располагаем кнопки посередине формы
    CB_OK.Left = (Me.Width / 2) - CB_OK.Width - 3
    CB_Cancel.Left = (Me.Width / 2) + 3
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "CB_Add_Click"
End Sub

Public Sub PS_GoToTimeColumn(ByVal index As Integer)
'Процедура перехода к колонке ВРЕМЕНИ с указанным индексом
'Если колонки с таким индексом нет - выход из проки
On Error GoTo EX
Dim TB As TextBox

    Set TB = Me.Controls("TB_Time_" & index)

    TB.SetFocus
    TB.SelStart = 0
    TB.SelLength = Len(TB.Text)
    
Exit Sub
EX:
'    Debug.Print "Нет такой колонки!"
    MsgBox "Такой колонки нет!"
    SaveLog Err, "PS_GoToTimeColumn"
End Sub
Public Sub PS_GoToDataColumn(ByVal index As Integer)
'Процедура перехода к колонке ДАННЫХ с указанным индексом
'Если колонки с таким индексом нет - выход из проки
On Error GoTo EX
Dim TB As TextBox

    Set TB = Me.Controls("TB_Data_" & index)
    
    TB.SetFocus
    TB.SelStart = 0
    TB.SelLength = Len(TB.Text)

Exit Sub
EX:
    MsgBox "Такой колонки нет!"
    SaveLog Err, "PS_GoToDataColumn"
End Sub

Public Function PF_CheckData() As Boolean
'Процедура возвращает ИСТИНА, если в форме нет ни одной записи с красным цветом
Dim ctrl As Control

    For Each ctrl In Me.Controls
        If ctrl.ForeColor = vbRed Then
            PF_CheckData = False
            Exit Function
        End If
    Next ctrl

PF_CheckData = True
End Function




Private Sub CB_Cancel_Click()
    '---Указываем, что нужно обновлять данные в графике
    RefreshNeed = False
    '---Закрываем форму
    CloseForm
End Sub

Private Sub CB_OK_Click()
    '---Проверяем, все ли данные указаны правильно
    If PF_CheckData = False Then
        MsgBox "Не все данные указаны правильно! Сохранение не возможно!", vbCritical
        Exit Sub
    End If
    
    '---Указываем, что нужно обновлять данные в графике
    RefreshNeed = True
    
    '---Закрываем форму
    CloseForm
End Sub

Private Sub UserForm_Terminate()
    '---Указываем, что НЕ нужно обновлять данные в графике
    RefreshNeed = False
    '---Закрываем форму
    CloseForm
End Sub

Private Sub CloseForm()
    ps_ClearTable
    Me.Hide
End Sub















Private Sub UserForm_Click()

End Sub




