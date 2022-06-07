VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Command 
   Caption         =   "Добавить команду"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frm_Command.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private shp As Visio.Shape
Private targetCellName As String

Public isOk As Boolean





Private Sub UserForm_Initialize()
    AddTimeFlag = False
End Sub

Private Sub UserForm_Activate()
'    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
'
'    Set shp = Application.ActiveWindow.Selection(1)
'
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & delimiter
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & " "
'    Me.txt_CommandText.Text = GetCurrentTime & delimiter
End Sub

Public Sub NewCommand(Optional ByRef shp_a As Visio.Shape = Nothing, Optional ByVal m As String = "2", Optional ByVal cmnd As String = "", Optional currentTime As Boolean = True)
    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
    
    If shp_a Is Nothing Then
        Set shp = Application.ActiveWindow.Selection(1)
    Else
        Set shp = shp_a
    End If
    
    targetCellName = ""
    
    'Если указано опция установки текущего времени, то ставим текущее время схемы, иначе - получаем последнее активное время выбранной фигуры
    If currentTime Then
        Me.txt_CommandText.text = GetCurrentTime & delimiter & m & delimiter & cmnd
    Else
        Me.txt_CommandText.text = getShpTime(shp) & delimiter & m & delimiter & cmnd
    End If
    
    Me.Show
End Sub

Public Sub CurrentCommand(ByRef shp_a As Visio.Shape, ByVal cellName As String)
    Set shp = shp_a
    targetCellName = cellName
    
    Me.txt_CommandText.text = shp.Cells(cellName).ResultStr(visUnitsString)
    
    Me.Show
End Sub

Private Sub btn_Ok_Click()
Dim rowI As Integer
Dim rowT As Integer
Dim rowA As Integer
Dim rowIName As String
Dim tagName As String
Dim frml As String
Dim targetCellNameShort As String
    
    'Если имени целевой ячейки нет, значит необходимо создавать новую ячейку, если есть - использовать ее для изменения
    If targetCellName = "" Then
'        rowIName = "GFS_Command_" & shp.RowCount(visSectionUser) + 1
        rowIName = "GFS_Command_" & GetNextNumber(shp)
        
        'строка в секции User
        rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
        shp.CellsSRC(visSectionUser, rowI, 0).Formula = """" & FixText(Me.txt_CommandText) & """"
        
        'строка в секции СмартТегов
        tagName = "Commands"
        If shp.CellExists("SmartTags.GFS_Commands", False) = 0 Then
            rowT = shp.AddNamedRow(visSectionSmartTag, "GFS_Commands", 0)
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagName).Formula = """" & tagName & """"
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagButtonFace).Formula = 346
        End If
        
        'строка в секции Action
        rowA = shp.AddNamedRow(visSectionAction, rowIName, 0)
        shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = """" & GetCommandText(75) & """"
        shp.CellsSRC(visSectionAction, rowA, visActionTagName).Formula = """" & tagName & """"
        frml = "CALLTHIS(" & Chr(34) & "RedactThisText" & _
                Chr(34) & "," & Chr(34) & "Управление_СиС" & Chr(34) & "," & _
                Chr(34) & "User." & rowIName & Chr(34) & ")"
        shp.CellsSRC(visSectionAction, rowA, visActionAction).FormulaU = frml
        
        'Добавляем строку проверки текущего времени User.CurrentDocTime, если ее нет
        If cellVal(shp, "User.CurrentDocTime", , "-1") < 0 Then
            rowIName = "CurrentDocTime"
            rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
            shp.CellsSRC(visSectionUser, rowI, 0).Formula = "TheDoc!User.CurrentTime"
            frml = "CALLTHIS(" & Chr(34) & "CheckEnd" & _
                Chr(34) & "," & Chr(34) & "Управление_СиС" & Chr(34) & _
                ",User.CurrentDocTime)"
            shp.CellsSRC(visSectionUser, rowI, 1).FormulaU = frml
        End If
        
        'Здесь в будущем можно будет указывать иконку самой строки команды
'        shp.CellsSRC(visSectionAction, rowA, visActionButtonFace).Formula = 346
               
    Else
        shp.Cells(targetCellName).Formula = """" & FixText(Me.txt_CommandText) & """"
        targetCellNameShort = Split(targetCellName, ".")(1)
        shp.Cells("Actions." & targetCellNameShort).Formula = """" & GetCommandText(75) & """"
    End If
    

    Me.isOk = True
    Me.Hide
End Sub
Private Sub btn_Cancel_Click()
    Me.isOk = False
    Me.Hide
End Sub

'Кнопка Отменить
Private Sub btn_Delete_Click()
Dim targetCellNameShort As String
Dim targetRowIndex As Integer
    
    ' Если целевой ячейки нет, значит, нужно просто отменить действие
    If targetCellName <> "" Then
        targetCellNameShort = Split(targetCellName, ".")(1)
        
        'Строка в секции User
        targetRowIndex = GetRowIndex("User." & targetCellNameShort)
        If targetRowIndex >= 0 Then
            shp.DeleteRow visSectionUser, targetRowIndex
        End If
        'Строка в секции Action
        targetRowIndex = GetRowIndex("Actions." & targetCellNameShort)   'Actions.GFS_Command_14
        If targetRowIndex >= 0 Then
            shp.DeleteRow visSectionAction, targetRowIndex
        End If
    End If
    
    'Пытаемся удалить смарт тег (для случая, если команд больше нет)
    TryDeleteSmartTag "Commands", "SmartTags.GFS_Commands"
    
    Me.Hide
End Sub

Private Function GetCurrentTime() As String
    GetCurrentTime = Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").ResultStr(visUnitsString)
End Function
Private Function FixText(ByVal str As String) As String
    FixText = Replace(str, Chr(34), "'")
End Function
Private Function GetCommandText(Optional ByVal l As Integer = 25) As String
Dim str As String

    On Error GoTo ex

    str = FixText(Me.txt_CommandText)
    If Len(str) < l Then
        GetCommandText = str
    Else
        GetCommandText = Left(str, l) & "..."
    End If

Exit Function
ex:
    GetCommandText = "***"
End Function
Private Function GetRowIndex(cellName As String) As Integer
    On Error GoTo ex
    GetRowIndex = shp.Cells(cellName).row
Exit Function
ex:
    GetRowIndex = -1
End Function

'Private Function GetShapeTime(ByRef shp As Visio.Shape) As String
'Dim shp As Visio.Shape
    
'    If Application.ActiveWindow.Selection.Count = 0 Then
'        Set shp = Application.ActiveWindow.Selection(1)
'        GetShapeTime = getShpTime(shp)
'    Else
'        GetShapeTime = GetCurrentTime
'        Exit Function
'    End If
'End Function

Private Sub TryDeleteSmartTag(stName As String, rowName As String)
'stName - название смарт-тега, rowName - название строки смарт-тега в секции SmartTags
Dim i As Integer
Dim smartTagRowIndex As Integer

    'Получаем инекс смарт тега и проверяем имеется ли такой смарт тег
    smartTagRowIndex = GetRowIndex(rowName)
    If smartTagRowIndex >= 0 Then
        'Проверяем все строки секции Actions на предмет наличия ссылок на указанный смарт тег
        For i = 0 To shp.RowCount(visSectionAction) - 1
            'Если есть хоть одна - выходим из процедуры не удаляя смарт тег
            If shp.CellsSRC(visSectionAction, i, visActionTagName).ResultStr(visUnitsString) = stName Then Exit Sub
        Next i

        shp.DeleteRow visSectionSmartTag, smartTagRowIndex
        'Удаляем так же и ячейку отслеживания времени
        On Error Resume Next
        shp.DeleteRow visSectionUser, shp.Cells("User.CurrentDocTime").row
    End If

End Sub

Public Function GetCommandTime() As Integer
    
    On Error GoTo ex
    GetCommandTime = Split(Me.txt_CommandText, delimiter)(1)
    
Exit Function
ex:
    GetCommandTime = 0
End Function

Public Function GetCommandDateTime() As Double
Dim tmp() As String
    On Error GoTo ex
'    GetCommandDateTime = Split(Me.txt_CommandText, delimiter)(1)  Split(Me.txt_CommandText, delimiter)(1)
    tmp = Split(Me.txt_CommandText, delimiter)
    GetCommandDateTime = DateAdd("n", tmp(1), CDate(tmp(0)))
    
Exit Function
ex:
    GetCommandDateTime = 0
End Function

Public Function GetCommandTextDate(ByVal cmnd_txt As String) As Date
' Получаем строку команды и возвращаем время окончания ее выполнения
Dim arr() As String
Dim tm As Date
Dim m As Integer

'    On Error GoTo ex
    
    arr = Split(cmnd_txt, delimiter)
    tm = arr(0)
    If IsNumeric(arr(1)) Then
        m = arr(1)
        GetCommandTextDate = DateAdd("n", m, tm)
    Else
        m = 0
        GetCommandTextDate = GetCurrentTime
    End If
    
'    GetCommandTextDate = DateAdd("n", m, tm)
    
Exit Function
ex:
    GetCommandTextDate = GetCurrentTime
End Function

Private Function getShpTime(ByRef shp As Visio.Shape) As String
'Возвращаем время окончания выполнения всех команд фигурой. Если команд нет, возвращаем время фигуры
Dim tm As Date
Dim tm_temp As Date
    
'    On Error GoTo ex
    
    tm = GetGFSShapeTime(shp)
    
    For i = 0 To shp.RowCount(visSectionUser) - 1
        If Left(shp.CellsSRC(visSectionUser, i, 0).Name, 17) = "User.GFS_Command_" Then
            tm_temp = GetCommandTextDate(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString))
            If tm_temp > tm Then tm = tm_temp
        End If
    Next i

getShpTime = tm
Exit Function
ex:
    getShpTime = GetCurrentTime
End Function

