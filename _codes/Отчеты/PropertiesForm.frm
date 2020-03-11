VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertiesForm 
   Caption         =   "Объект пожара"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   OleObjectBlob   =   "PropertiesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PropertiesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CBut_OK_Click()
Me.Hide
sp_DataRefresh
End Sub

Private Sub UserForm_Activate()
'Открытие формы для просмотра и редактирования свойств объекта пожара
'---Объявляем переменные
Dim vpVS_DocShape As Visio.Shape

'---Инициируем объект Шэйп-листа документа
    Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---Получаем данные из Шэйп-листа документа
    Me.TB_City = vpVS_DocShape.Cells("User.City").ResultStr(Visio.visNone)
    Me.TB_Adress = vpVS_DocShape.Cells("User.Adress").ResultStr(Visio.visNone)
    Me.CB_FireRating = vpVS_DocShape.Cells("User.FireRating").ResultStr(Visio.visNone)
    Me.CB_Object = vpVS_DocShape.Cells("User.Object").ResultStr(Visio.visNone)

End Sub



Private Sub UserForm_Initialize()
'---Заполняем списки
'---Объявляем переменные
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim i As Integer

    On Error GoTo EX
    '---Список степеней огнестойкости
    For i = 1 To 5
        Me.CB_FireRating.AddItem (i)
    Next i
    
    '---Список степеней огнестойкости
    '---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT Описание, [Категория] " & _
        "FROM З_Интенсивности " & _
        "WHERE (([Категория])='Здания и сооружения')" & _
        "ORDER BY З_Интенсивности.Описание;"

    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        
    '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
    With rst
        .MoveFirst
        Do Until .EOF
            Me.CB_Object.AddItem (![Описание])
            .MoveNext
        Loop
    End With

Exit Sub
EX:
    SaveLog Err, "UserForm_Initialize"
End Sub

Private Sub UserForm_Terminate()
'Скрытие формы и обновление данных
sp_DataRefresh
End Sub

Private Sub sp_DataRefresh()
'---Объявляем переменные
Dim vpVS_DocShape As Visio.Shape

'---Инициируем объект Шэйп-листа документа
Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---Получаем данные из Шэйп-листа документа
vpVS_DocShape.Cells("User.City").FormulaU = Chr(34) & Me.TB_City.value & Chr(34)
vpVS_DocShape.Cells("User.Adress").FormulaU = Chr(34) & Me.TB_Adress.value & Chr(34)
vpVS_DocShape.Cells("User.FireRating").FormulaU = Chr(34) & Me.CB_FireRating.value & Chr(34)
vpVS_DocShape.Cells("User.Object").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
End Sub
