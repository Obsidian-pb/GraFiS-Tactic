VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Sequencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Класс для автоматизации добавления основных сведений о фигурах ГраФиС при их добавлении
' Общая логика: при размещении новых фигур ГраФиС поверх имеющихся им передаются сведения имеющихся (Подразделение, Позывной, Время)


Public WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1

Private t As String






' Процедуры класса
Private Sub Class_Initialize()
    Set app = Visio.Application
    Debug.Print ("sequencer activateed")
End Sub
Private Sub Class_Terminate()
    Set app = Nothing
    Debug.Print ("sequencer DEactivateed")
End Sub



' Процедуры функционала приложения
'Private Sub app_AppActivated(ByVal app As IVApplication)
'
'End Sub
Private Sub app_ShapeAdded(ByVal shp As IVShape)
Dim mainShp As Visio.Shape
    On Error GoTo ex
'    DoEvents
    If IsGFSShape(shp) Then
        Set mainShp = GetGraFisShape(shp)
        If Not mainShp Is Nothing Then
            '---Добавляем команду
            CheckAndAddCommand mainShp, shp
            
            GetShpData mainShp, shp
            
'            frm_Command.NewCommand
            
        End If
        
'        Debug.Print mainShp.Name
    End If

Exit Sub
ex:
    Debug.Print ("None")
End Sub

Private Sub CheckAndAddCommand(ByRef parent As Visio.Shape, ByRef child As Visio.Shape)
'Dim t As String
    
    If IsGFSShapeWithIP(parent, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)) Then
        Select Case cellval(child, "user.IndexPers")
            Case Is = indexPers.ipStvolRuch
                t = AskTimeDiff
                frm_Command.NewCommand t, "Подать ствол "
            Case Is = indexPers.ipGDZSZvenoVozd
                t = AskTimeDiff("1")
                frm_Command.NewCommand t, "Сформировать звено ГДЗС"
            Case Is = indexPers.ipGDZSPost
                t = AskTimeDiff("1")
                frm_Command.NewCommand t, "Выставить пост безопасности ГДЗС"
            Case Is = indexPers.ipLestn3Kol
                t = AskTimeDiff("2")
                frm_Command.NewCommand t, "Установить трехколенную лестницу"
            Case Is = indexPers.ipLestnPalka
                t = AskTimeDiff("1")
                frm_Command.NewCommand t, "Установить лестницу-палку"
            Case Is = indexPers.ipLestnSturm
                t = AskTimeDiff("1")
                frm_Command.NewCommand t, "Установить штурмовую лестницу"
            Case Is = indexPers.ipRazvetvlenie
                t = AskTimeDiff("3")
                frm_Command.NewCommand t, "Проложить магистральную рукавную линию, установить разветвление"
            Case Else
                t = "0"
        End Select
    Else
        t = "0"
    End If
    
    
End Sub

Private Function AskTimeDiff(Optional ByVal d As String = "2") As String
'Запрашиваем у пользователя время требуемое на выполнение команды
    AskTimeDiff = InputBox("Сколько времени требуется на выполнение данной команды?", , d)
End Function


' Служебные функции
Private Function GetGraFisShape(shp As Visio.Shape) As Visio.Shape
' Ищем фигуру на которую была наложена текущая (shp) и если эта фигура является фигурой ГраФиС, возвращаем ее
Dim tmpShp As Visio.Shape
Dim sel As Visio.Selection
    
    Set sel = Application.ActivePage.SpatialSearch(cellval(shp, "PinX", visInches), cellval(shp, "PinY", visInches), _
                                                    VisSpatialRelationCodes.visSpatialTouching + VisSpatialRelationCodes.visSpatialContainedIn, _
                                                    0, _
                                                    visSpatialFrontToBack + visSpatialIgnoreVisible)
    For Each tmpShp In sel
        If Not (tmpShp.ID = shp.ID) Then
            ' if shape is grafis shape
            If IsGFSShape(shp) Then
                ' if not fire or fog
                If Not IsGFSShapeWithIP(shp, Array(64, 71)) Then
                    Set GetGraFisShape = tmpShp
                    Exit Function
                End If
            End If
        End If
    Next tmpShp
    
Set GetGraFisShape = Nothing
End Function

Private Function GetShpData(ByRef mainShp As Visio.Shape, shp As Visio.Shape) As Variant
Dim tmpDateTime As Double
Dim tmpDateTimeS As String
    On Error Resume Next
    
'    tmpDateTime = cellval(mainShp, "Prop.SetTime", visDate) + _
'                    cellval(mainShp, "Prop.ArrivalTime", visDate) + _
'                    cellval(mainShp, "Prop.FormingTime", visDate)
'    tmpDateTime = DateAdd("n", Int(t), tmpDateTime)
    
    tmpDateTime = GetGFSShapeTime(mainShp)
    tmpDateTime = DateAdd("n", Int(t), tmpDateTime)
                    
    If tmpDateTime > 0 Then
        SetCellVal shp, "Prop.Unit", cellval(mainShp, "Prop.Unit", visUnitsString)
        SetCellVal shp, "Prop.Call", cellval(mainShp, "Prop.Call", visUnitsString)
    
    
        'Временно отключил, т.к. выяснилось, что на правктике удобнее работать с текущим временем схемы
        If shp.CellExists("User.InPage", 0) = 0 Then
            shp.AddNamedRow visSectionUser, "InPage", visRowUser
            SetCellVal shp, "User.InPage", 1
            SetCellVal shp, "User.InPage.Prompt", "+"
        End If
    'Временно отключил, т.к. выяснилось, что на правктике удобнее работать с текущим временем схемы
        tmpDateTimeS = ToTimeString(tmpDateTime)
    '    SetCellVal shp, "Prop.SetTime", tmpDateTimeS
    '    SetCellVal shp, "Prop.FormingTime", tmpDateTimeS
        SetCellFrml shp, "Prop.SetTime", tmpDateTimeS
        SetCellFrml shp, "Prop.FormingTime", tmpDateTimeS
    End If
    
End Function

Private Function GetGFSShapeTime(ByRef shp As Visio.Shape) As Double
    
    GetGFSShapeTime = cellval(shp, "Prop.SetTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellval(shp, "Prop.FormingTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellval(shp, "Prop.ArrivalTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function

GetGFSShapeTime = 0
End Function

Private Function ToTimeString(ByVal dblVal As Double) As String
Dim tmp As String
    tmp = Replace(str(dblVal), ",", ".")
    ToTimeString = "DATETIME(" & tmp & ")"
End Function
