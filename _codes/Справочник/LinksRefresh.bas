Attribute VB_Name = "LinksRefresh"
Option Compare Database

Private Sub s_TableLinksRefresh(ass_BaseName As String)
'Процедура перелинковки всех таблиц

'---Объявляем переменные
Dim db As DAO.Database
Dim td As DAO.TableDef
Dim vss_CurPath As String
Dim vss_ConnectionString As String
Dim vss_ConnectionStringAlter As String
Dim vsO_tdfsCur As DAO.TableDefs

'---Присваиваем необходимые для работы переменные
vss_CurPath = Application.CurrentProject.path & "\" & "Signs.fdb"
vss_ConnectionString = ";DATABASE=" & vss_CurPath
Set db = CurrentDb
Set vsO_tdfsCur = db.TableDefs

''---Устанавливаем альтернативную строку подключения для БД SRVC.mdb - нужна для таблицы подразделения
'vss_CurPath = Application.CurrentProject.Path & "\" & "Srvc.mdb"
'vss_ConnectionStringAlter = ";DATABASE=" & vss_CurPath

'---Запускаем цикл прилинковки
For Each td In db.TableDefs
    If td.Attributes = dbAttachedTable Then
        td.Connect = vss_ConnectionString 'Прилинковываем к новому местоположению
        td.RefreshLink
    End If
Next td

'---Закрываем соединение с базой данных
vsO_tdfsCur.Refresh
db.Close
End Sub


Public Function PF_LinkCheck()
'Функция проверки работоспособности текущей прилинковки таблиц БД

'---Объвяляем переменные
Dim vo_db_CurrentDataBase As DAO.Database
Dim vs_NeededLink As String

'---Присваиваем переменные необходимые для работы
Set vo_db_CurrentDataBase = CurrentDb
vs_NeededLink = ";DATABASE=" & Application.CurrentProject.path & "\Signs.fdb"

'---Проверяем соответствует ли строка подключения таблицы "АА" иекущему распооложения связанной базы данных (SignsNew.mdb)
If Not vs_NeededLink = vo_db_CurrentDataBase.TableDefs("АА").Connect Then
    s_TableLinksRefresh ("SignsNew.mdb") 'Если не соответствует, то перпелинковываем все пользовательские таблицы
End If

'---Закрываем связь с базой данных
vo_db_CurrentDataBase.Close
End Function



