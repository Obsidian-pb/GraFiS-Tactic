Attribute VB_Name = "m_Roads"
'-----------------Модуль для процедур работы с дорогами----------------------------
Option Explicit



Sub CheckConnectionBeg(ShpObj As Visio.Shape)
'Процедура проверяет не подключена ли к началу дороги какая-либо другая дорога
'И если подключена - определяет положение поребриков
Dim i As Integer
Dim Conn As Visio.Connect
Dim vO_ToSheet As Visio.Shape
Dim vO_Formula As String
Dim vB_BeginConnectionHave As Boolean
Dim vB_EndConnectionHave As Boolean
    
    On Error GoTo EX
    
    vB_BeginConnectionHave = False
    vB_EndConnectionHave = False
    
    For Each Conn In ShpObj.Connects
        
'        Debug.Print conn.FromCell.Name & " - " & Left(conn.ToCell.Name, 11)
'        Debug.Print conn.FromCell.Name & " - " & conn.ToCell.Name
'        Debug.Print conn.ToSheet.Name
'        Debug.Print conn.ToSheet.Cells("User.IndexPers").ResultInt(visNumber, 0)
'        Debug.Print ShpObj.Connects.Count
'        Exit Sub  ' Временно!!!!!!!
        
        '---Проверяем к какой фигуре осуществлен коннект - к проезжей части или к центральной линии
        'ПРОЕЗЖАЯ ЧАСТЬ
        If Conn.ToSheet.Cells("User.IndexPers") = 130 Then
            Set vO_ToSheet = Conn.ToSheet
            
            If Conn.FromCell.Name = "EndX" Then
                '---Привязка
                '---проверяем к какому концу проезжей части фигура присоединена
                If Conn.ToCell.Name = "Connections.X1" Then
                    vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle)"
                ElseIf Conn.ToCell.Name = "Connections.X2" Then
                    vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle-180deg)"
                End If
                ShpObj.Cells("User.AngleWithEndShape").FormulaU = vO_Formula
                '---Ширина фигуры к которой приклеено
                vO_Formula = vO_ToSheet.Name & "!Height"
                ShpObj.Cells("User.HeightEndShape").FormulaU = vO_Formula
                '---Отмечаем, что к конечной точке присоединена фигура
                vB_EndConnectionHave = True
                '---Скрываем конечное скругление и поребрик
                ShpObj.Cells("Actions.ShowEndRound.Checked") = 0
'                ShpObj.Cells("Actions.ShowEndArc.Checked") = 0
                
                '---Перемещаем фигуру вперед
                ShpObj.BringToFront
                
                '---Указываем, что соединение не с торцами другой дороги
                ShpObj.Cells("User.EndIsConnect") = 1
                
            ElseIf Conn.FromCell.Name = "BeginX" Then
                '---Привязка
                '---проверяем к какому концу проезжей части фигура присоединена
                If Conn.ToCell.Name = "Connections.X1" Then
                    vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle-180deg)"
                ElseIf Conn.ToCell.Name = "Connections.X2" Then
                    vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle)"
                End If
                ShpObj.Cells("User.AngleWithBeginShape").FormulaU = vO_Formula
                '---Ширина фигуры к которой приклеено
                vO_Formula = vO_ToSheet.Name & "!Height"
                ShpObj.Cells("User.HeightBeginShape").FormulaU = vO_Formula
                '---Отмечаем, что к начальной точке присоединена фигура
                vB_BeginConnectionHave = True
                '---Скрываем начальное скругление и поребрик
                ShpObj.Cells("Actions.ShowBeginRound.Checked") = 0
'                ShpObj.Cells("Actions.ShowBeginArc.Checked") = 0
                
                '---Перемещаем фигуру вперед
                ShpObj.BringToFront
                
                '---Указываем, что соединение не с торцами другой дороги
                ShpObj.Cells("User.BeginIsConnect") = 1
            End If
                
        
        'ЦЕНТРАЛЬНАЯ ЛИНИЯ
        ElseIf Conn.ToSheet.Cells("User.IndexPers") = 131 Then
            Set vO_ToSheet = Conn.ToSheet.Parent
            
            If Conn.FromCell.Name = "EndX" And Left(Conn.ToCell.Name, 11) = "Connections" Then
                '---Привязка
                vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle)"
                ShpObj.Cells("User.AngleWithEndShape").FormulaU = vO_Formula
                '---Ширина фигуры к которой приклеено
                vO_Formula = vO_ToSheet.Name & "!Height"
                ShpObj.Cells("User.HeightEndShape").FormulaU = vO_Formula
                '---Отмечаем, что к конечной точке присоединена фигура
                vB_EndConnectionHave = True
                '---Скрываем конечное скругление и поребрик
                ShpObj.Cells("Actions.ShowEndRound.Checked") = 0
'                ShpObj.Cells("Actions.ShowEndArc.Checked") = 0
                
                
                '---Перемещаем фигуру вперед
                ShpObj.BringToFront
                
                '---Указываем, что соединение с торцами другой дороги
                ShpObj.Cells("User.EndIsConnect") = 0
                
            ElseIf Conn.FromCell.Name = "BeginX" And Left(Conn.ToCell.Name, 11) = "Connections" Then
                '---Привязка
                vO_Formula = "ANG360(User.UsedAngle-" & vO_ToSheet.Name & "!User.UsedAngle)"
                ShpObj.Cells("User.AngleWithBeginShape").FormulaU = vO_Formula
                '---Ширина фигуры к которой приклеено
                vO_Formula = vO_ToSheet.Name & "!Height"
                ShpObj.Cells("User.HeightBeginShape").FormulaU = vO_Formula
                '---Отмечаем, что к конечной точке присоединена фигура
                vB_BeginConnectionHave = True
                '---Скрываем начальное скругление и поребрик
                ShpObj.Cells("Actions.ShowBeginRound.Checked") = 0
'                ShpObj.Cells("Actions.ShowBeginArc.Checked") = 0
                
                '---Перемещаем фигуру вперед
                ShpObj.BringToFront
                
                '---Указываем, что соединение с торцами другой дороги
                ShpObj.Cells("User.BeginIsConnect") = 0
            End If
            
        End If

        '---Проверка соединения торцов
        
    Next Conn
    
    '---В случае, если соединений с конца или начала фигуры нет, очищаем формулы для соответствующих ячеек начала и конца
    If vB_BeginConnectionHave = False Then
        ShpObj.Cells("User.AngleWithBeginShape").FormulaU = "90deg"
        ShpObj.Cells("User.HeightBeginShape").FormulaU = "0"
        '---Показываем начальное скругление и поребрик
        ShpObj.Cells("Actions.ShowBeginRound.Checked") = 1
'        ShpObj.Cells("Actions.ShowBeginArc.Checked") = 1
    End If
    If vB_EndConnectionHave = False Then
        ShpObj.Cells("User.AngleWithEndShape").FormulaU = "90deg"
        ShpObj.Cells("User.HeightEndShape").FormulaU = "0"
        '---Показываем конечное скругление и поребрик
        ShpObj.Cells("Actions.ShowEndRound.Checked") = 1
'        ShpObj.Cells("Actions.ShowEndArc.Checked") = 1
    End If

    '---Перемещаем все дороги назад
    SendAllRoadsBack ShpObj

Exit Sub
EX:
    SendAllRoadsBack ShpObj
    SaveLog Err, "CheckConnectionBeg"
End Sub


Private Sub SendAllRoadsBack(ShpObj As Visio.Shape)
'Прока отправляет все дороги на задний фон
Dim vsoSelection As Visio.Selection
    
    On Error GoTo EX
    
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Проезжая часть")
    Application.ActiveWindow.Selection = vsoSelection

    Application.ActiveWindow.Selection.SendToBack
    ActiveWindow.DeselectAll
    
    Application.ActiveWindow.Select ShpObj, visSelect

Exit Sub
EX:
    SaveLog Err, "SendAllRoadsBack"
End Sub







