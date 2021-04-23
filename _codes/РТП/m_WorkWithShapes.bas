Attribute VB_Name = "m_WorkWithShapes"
Option Explicit

'Перечисление возможных состояний тактической единицы
Public Enum tactState
    tsInProgress = 0
    tsWaiting = 1
    tsEnd = 2
    tsError = 3
    tsNotStarted = 4
    tsUnknown = -1
End Enum

Public Const delimiter = " | "





Public Sub RedactThisText(ByRef shp As Visio.Shape, ByVal cellName As String)
    frm_Command.CurrentCommand shp, cellName
End Sub

Public Sub CheckEnd(ByRef shp As Visio.Shape, ByVal t As Date)
    
    Select Case GetTactState(shp)
        Case Is = tactState.tsInProgress
            SetCellFrml shp, "SmartTags.GFS_Commands.ButtonFace", 346
            SetCellFrml shp, "SmartTags.GFS_Commands.Description", "Выполняет поставленные задачи"
        Case Is = tactState.tsWaiting
            SetCellFrml shp, "SmartTags.GFS_Commands.ButtonFace", 1089
            SetCellFrml shp, "SmartTags.GFS_Commands.Description", "Ожидает команд"
        Case Is = tactState.tsEnd
            SetCellFrml shp, "SmartTags.GFS_Commands.ButtonFace", 840
            SetCellFrml shp, "SmartTags.GFS_Commands.Description", "Закончил работу на пожаре (убыл)"
        Case Is = tactState.tsNotStarted
            SetCellFrml shp, "SmartTags.GFS_Commands.ButtonFace", 2743
            SetCellFrml shp, "SmartTags.GFS_Commands.Description", "Выполнение не началось"
        Case Is = tactState.tsError
            SetCellFrml shp, "SmartTags.GFS_Commands.ButtonFace", 463
            SetCellFrml shp, "SmartTags.GFS_Commands.Description", "ОШИБКА - проверьте корректность данных"
    End Select

End Sub

Private Function GetTactState(ByRef shp As Visio.Shape) As tactState
Dim i As Integer
Dim curTime As Date
Dim curCommandState As tactState
Dim firstState As tactState
    
    firstState = tactState.tsUnknown
    
    For i = 0 To shp.RowCount(visSectionAction) - 1
        If left(shp.CellsSRC(visSectionAction, i, 0).rowName, 11) = "GFS_Command" Then
            curCommandState = GetCommandState(cellVal(shp, "User.CurrentDocTime", visDate), _
                           shp.CellsSRC(visSectionAction, i, visActionMenu).ResultStr(visUnitsString))
            Select Case curCommandState
                Case Is = tactState.tsInProgress
                    shp.CellsSRC(visSectionAction, i, visActionButtonFace).Formula = 346
                Case Is = tactState.tsWaiting
                    shp.CellsSRC(visSectionAction, i, visActionButtonFace).Formula = 837
                Case Is = tactState.tsEnd
                    shp.CellsSRC(visSectionAction, i, visActionButtonFace).Formula = 837
                Case Is = tactState.tsError
                    shp.CellsSRC(visSectionAction, i, visActionButtonFace).Formula = 463
                Case Is = tactState.tsNotStarted
                    shp.CellsSRC(visSectionAction, i, visActionButtonFace).Formula = 2743
            End Select
            
            'Запоминаем состояние первой команды
            If firstState = tactState.tsUnknown Then
                firstState = curCommandState
            End If
        End If
    Next i
    
    'Проверяем начала ли выполняться первая команда. Если да, то возвращае обычное значение, если нет - то состояние "Не началось"
    If firstState = tsNotStarted Then
        GetTactState = tactState.tsNotStarted
    Else
        If curCommandState = tsNotStarted Then
            GetTactState = tsInProgress
        Else
            GetTactState = curCommandState
        End If
    End If
    
End Function

Public Function GetCommandState(ByRef curTime As Date, ByVal commandText As String) As tactState
Dim comArr() As String
Dim comArrLen As Integer
Dim startTime As Date
Dim durationS As String
Dim durationI As Integer

    On Error GoTo EX

    'Если время выполнения не указано, работа не имеет ограничения
    'Если время выполнения = *, единица закончила работу на пожаре
    'Если время выполнения = число, прибавляем
    comArr = Split(commandText, delimiter)
    comArrLen = UBound(comArr)
    startTime = CDate(comArr(0))
    durationS = Trim(comArr(1))
    
    curTime = DateAdd("s", 10, curTime)
    
    If curTime < startTime Then
        GetCommandState = tsNotStarted
    Else
        If durationS = "*" Then
'            If curTime >= startTime Then
                GetCommandState = tsEnd
'            Else
'                GetCommandState = tsInProgress
'            End If
        ElseIf durationS = "" Or durationS = " " Then
            GetCommandState = tsInProgress
        Else
            durationI = Int(durationS)
            If curTime >= DateAdd("n", durationI, startTime) Then
                GetCommandState = tsWaiting
            Else
                GetCommandState = tsInProgress
            End If
        End If
    End If
    
Exit Function
EX:
    GetCommandState = tsError
End Function
