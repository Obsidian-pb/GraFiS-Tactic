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
Const delimiter = " | "








Private Sub UserForm_Activate()
'    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
'
'    Set shp = Application.ActiveWindow.Selection(1)
'
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & delimiter
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & " "
'    Me.txt_CommandText.Text = GetCurrentTime & delimiter
End Sub

Public Sub NewCommand()
    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    targetCellName = ""
    
    Me.txt_CommandText.Text = GetCurrentTime & delimiter
    
    Me.Show
End Sub

Public Sub CurrentCommand(ByRef shp_a As Visio.Shape, ByVal cellName As String)
    Set shp = shp_a
    targetCellName = cellName
    
    Me.txt_CommandText.Text = shp.Cells(cellName).ResultStr(visUnitsStig)
    
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
        rowIName = "GFS_Command_" & shp.RowCount(visSectionUser) + 1
        
        'строка в секции User
        rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
        shp.CellsSRC(visSectionUser, rowI, 0).Formula = """" & FixText(Me.txt_CommandText) & """"
        
        'строка в секции СмартТегов
        tagName = "Commands"
        If shp.CellExists("SmartTags.GFS_Commands", False) = 0 Then
            rowT = shp.AddNamedRow(visSectionSmartTag, "GFS_Commands", 0)
    '        shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagDescription).Formula = "User." & rowIName
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagName).Formula = """" & tagName & """"
        End If
        
        'строка в секции Action
'        rowA = shp.AddNamedRow(visSectionAction, "GFS_Command_" & shp.RowCount(visSectionAction) + 1, 0)
        rowA = shp.AddNamedRow(visSectionAction, rowIName, 0)
        shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = """" & GetCommandText(75) & """"
'        shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = "User." & rowIName
        shp.CellsSRC(visSectionAction, rowA, visActionTagName).Formula = """" & tagName & """"
        frml = "CALLTHIS(" & Chr(34) & "RedactThisText" & _
                Chr(34) & "," & Chr(34) & "РТП" & Chr(34) & "," & _
                Chr(34) & "User." & rowIName & Chr(34) & ")"
        shp.CellsSRC(visSectionAction, rowA, visActionAction).FormulaU = frml
    Else
        shp.Cells(targetCellName).Formula = """" & FixText(Me.txt_CommandText) & """"
        targetCellNameShort = Split(targetCellName, ".")(1)
        shp.Cells("Actions." & targetCellNameShort).Formula = """" & GetCommandText(75) & """"
    End If
    

    
    Me.Hide
End Sub
Private Sub btn_Cancel_Click()
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
'Dim comarr() As String
    On Error GoTo ex
'    comarr = Split(Me.txt_CommandText, delimiter)
'    str = comarr(UBound(comarr))
    str = FixText(Me.txt_CommandText)
    If Len(str) < l Then
        GetCommandText = str
    Else
        GetCommandText = Left(str, l) & "..."
    End If
'    Debug.Print GetCommandText
Exit Function
ex:
    GetCommandText = "***"
End Function
Private Function GetRowIndex(cellName As String) As Integer
    On Error GoTo ex
    GetRowIndex = shp.Cells(cellName).Row
Exit Function
ex:
    GetRowIndex = -1
End Function

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
    End If

End Sub
