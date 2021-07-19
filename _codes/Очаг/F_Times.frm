VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Times 
   Caption         =   "Временные показатели модели"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8505.001
   OleObjectBlob   =   "F_Times.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Times"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Enum TimeProp
Const fir = "Prop.FireTime"
Const fin = "Prop.FindTime"
Const inf = "Prop.InfoTime"
Const fArr = "Prop.FirstArrivalTime"
Const fStv = "Prop.FirstStvolTime"
Const loc = "Prop.LocalizationTime"
Const log = "Prop.LOGTime"
Const lpp = "Prop.LPPTime"
Const endF = "Prop.FireEndTime"
'End TimeProp
''Enum ControlType
'Const bmd_CT = ""
'
''End ControlType

Public formShp As Visio.Shape


Public Sub ShowMe(ByRef shp As Visio.Shape)
    Set formShp = shp
    LoadTimes
    Me.Show
End Sub

'---------------Основные кнопки формы------------------
Private Sub btn_CLose_Click()
    Me.Hide
End Sub

Private Sub btn_OK_Click()
    SaveTimes
End Sub
Private Sub btn_OKCLose_Click()
    SaveTimes.Hide
End Sub

'---------------Кнопки формы для манипуляций с данными------------------
    '---------------Очистка даты----------------------------------------
    Private Sub btn_FindTime_Clear_Click()
        ClearDate fin
    End Sub
    Private Sub btn_FireEndTime_Clear_Click()
        ClearDate endF
    End Sub
    Private Sub btn_FireTime_Clear_Click()
        ClearDate fir
    End Sub
    Private Sub btn_FirstArrivalTime_Clear_Click()
        ClearDate fArr
    End Sub
    Private Sub btn_FirstStvolTime_Clear_Click()
        ClearDate fStv
    End Sub
    Private Sub btn_InfoTime_Clear_Click()
        ClearDate inf
    End Sub
    Private Sub btn_LocalizationTime_Clear_Click()
        ClearDate loc
    End Sub
    Private Sub btn_LOGTime_Clear_Click()
        ClearDate log
    End Sub
    Private Sub btn_LPPTime_Clear_Click()
        ClearDate lpp
    End Sub
    
    '---------------Добавление 1 минуты----------------------------------------
    Private Sub btn_FindTime_PlusMin_Click()
        TimeAdd fin
    End Sub
    Private Sub btn_FireEndTime_PlusMin_Click()
        TimeAdd endF
    End Sub
    Private Sub btn_FireTime_PlusMin_Click()
        TimeAdd fir
    End Sub
    Private Sub btn_FirstArrivalTime_PlusMin_Click()
        TimeAdd fArr
    End Sub
    Private Sub btn_FirstStvolTime_PlusMin_Click()
        TimeAdd fStv
    End Sub
    Private Sub btn_InfoTime_PlusMin_Click()
        TimeAdd inf
    End Sub
    Private Sub btn_LocalizationTime_PlusMin_Click()
        TimeAdd loc
    End Sub
    Private Sub btn_LOGTime_PlusMin_Click()
        TimeAdd log
    End Sub
    Private Sub btn_LPPTime_PlusMin_Click()
        TimeAdd lpp
    End Sub
    
    '---------------Удаление 1 минуты----------------------------------------
    Private Sub btn_FindTime_DelMin_Click()
        TimeAdd fin, -1
    End Sub
    Private Sub btn_FireEndTime_DelMin_Click()
        TimeAdd endF, -1
    End Sub
    Private Sub btn_FireTime_DelMin_Click()
        TimeAdd fir, -1
    End Sub
    Private Sub btn_FirstArrivalTime_DelMin_Click()
        TimeAdd fArr, -1
    End Sub
    Private Sub btn_FirstStvolTime_DelMin_Click()
        TimeAdd fStv, -1
    End Sub
    Private Sub btn_InfoTime_DelMin_Click()
        TimeAdd inf, -1
    End Sub
    Private Sub btn_LocalizationTime_DelMin_Click()
        TimeAdd loc, -1
    End Sub
    Private Sub btn_LOGTime_DelMin_Click()
        TimeAdd log, -1
    End Sub
    Private Sub btn_LPPTime_DelMin_Click()
        TimeAdd lpp, -1
    End Sub
    
    '-----------Установка времен---------------------------
    Private Sub btn_FireTime_Cur_Click()
        SetVal fir, CStr(Now)
    End Sub
    Private Sub btn_FindTime_Prev_Click()
        SetVal fin, GetVal(fir)
    End Sub
    Private Sub btn_InfoTime_Prev_Click()
        SetVal inf, GetVal(fin)
    End Sub
    Private Sub btn_FirstArrivalTime_Prev_Click()
        SetVal fArr, GetVal(inf)
    End Sub
    Private Sub btn_FirstStvolTime_Prev_Click()
        SetVal fStv, GetVal(fArr)
    End Sub
    Private Sub btn_LocalizationTime_Prev_Click()
        SetVal loc, GetVal(fStv)
    End Sub
    Private Sub btn_LOGTime_Prev_Click()
        SetVal log, GetVal(loc)
    End Sub
    Private Sub btn_LPPTime_Prev_Click()
        SetVal lpp, GetVal(log)
    End Sub
    Private Sub btn_FireEndTime_Prev_Click()
        SetVal endF, GetVal(lpp)
    End Sub

    '-------------Анализ данных--------------
    Private Sub btn_FirstArrivalTime_Calc_Click()
    Dim dt As Date
        dt = GetMinArrTime
        If dt = 0 Then
            MsgBox "На странице нет фигур со свойством 'Время прибытия'."
        Else
            SetVal fArr, CStr(dt)
        End If
    End Sub
    Private Sub btn_FirstStvolTime_Calc_Click()
    Dim dt As Date
        dt = GetMinStvTime
        If dt = 0 Then
            MsgBox "На странице нет фигур со свойством 'Время подачи'."
        Else
            SetVal fStv, CStr(dt)
        End If
    End Sub

    '-------------Определение разниц между временами--------------
    Private Sub txt_FindTime_Change()
        CheckDiffs
        CheckField fin
    End Sub
    Private Sub txt_FireEndTime_Change()
        CheckDiffs
        CheckField endF
    End Sub
    Private Sub txt_FireTime_Change()
        CheckDiffs
        CheckField fir
    End Sub
    Private Sub txt_FirstArrivalTime_Change()
        CheckDiffs
        CheckField fArr
    End Sub
    Private Sub txt_FirstStvolTime_Change()
        CheckDiffs
        CheckField fStv
    End Sub
    Private Sub txt_InfoTime_Change()
        CheckDiffs
        CheckField inf
    End Sub
    Private Sub txt_LocalizationTime_Change()
        CheckDiffs
        CheckField loc
    End Sub
    Private Sub txt_LOGTime_Change()
        CheckDiffs
        CheckField log
    End Sub
    Private Sub txt_LPPTime_Change()
        CheckDiffs
        CheckField lpp
    End Sub
    '-------------Вставка текущих значений по двойному клику--------------
    Private Sub txt_FireTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal fir, CStr(Now)
    End Sub
    Private Sub txt_FindTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal fin, CStr(Now)
    End Sub
    Private Sub txt_FirstArrivalTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal fArr, CStr(Now)
    End Sub
    Private Sub txt_FirstStvolTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal fStv, CStr(Now)
    End Sub
    Private Sub txt_InfoTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal inf, CStr(Now)
    End Sub
    Private Sub txt_LocalizationTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal loc, CStr(Now)
    End Sub
    Private Sub txt_LOGTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal log, CStr(Now)
    End Sub
    Private Sub txt_LPPTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal lpp, CStr(Now)
    End Sub
    Private Sub txt_FireEndTime_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        SetVal endF, CStr(Now)
    End Sub
'    '------------Добавление строк в лист страницы
'    Private Sub cb_ShowInPage_Click()
'        AddPageTimeProps
'    End Sub


'---------------Функции формы--------------------------
Private Function GetControl(ByVal tag As String) As Object
Dim obj As Object
    
    For Each obj In Controls
        If obj.tag = tag Then
            Set GetControl = obj
            Exit Function
        End If
    Next obj
    
Set GetControl = Nothing
End Function

Private Sub SetVal(ByVal tag As String, ByVal val As Variant)
Dim obj As Object
    Set obj = GetControl(tag)
    obj.Text = val
End Sub
Private Sub SetCaption(ByVal tag As String, ByVal val As Variant)
Dim obj As Object
    Set obj = GetControl(tag)
    obj.Caption = val
End Sub

Private Sub SetColor(ByVal tag As String, ByVal col As Long)
Dim obj As Object
    Set obj = GetControl(tag)
    obj.ForeColor = col
End Sub

Private Function GetVal(ByVal tag As String) As Variant
Dim obj As Object
    Set obj = GetControl(tag)
    GetVal = obj.value
End Function

Private Function GetTimeDiff(ByVal tag1 As String, ByVal tag2 As String) As Integer
Dim val1 As Date
Dim val2 As Date

    GetTimeDiff = DateDiff("n", GetVal(tag1), GetVal(tag2))
End Function

Private Function isCorrectData(ByVal txt As String) As Boolean
Dim test As Date
    On Error GoTo EX
    test = CDate(txt)
    
isCorrectData = True
Exit Function
EX:
    isCorrectData = False
End Function

Private Sub CheckDiffs()
    CheckDiff fir, fin
    CheckDiff fin, inf
    CheckDiff inf, fArr
    CheckDiff fArr, fStv
    CheckDiff fStv, loc
    CheckDiff loc, log
    CheckDiff log, lpp
    CheckDiff lpp, endF
End Sub
Private Sub CheckDiff(ByVal tag1 As String, ByVal tag2 As String)
Dim dif As Integer
Dim lblTag As String
    
    On Error Resume Next
    
    lblTag = tag2 & ".lbl"
    
    dif = GetTimeDiff(tag1, tag2)
    SetCaption lblTag, dif
    If dif < 0 Then
        SetColor lblTag, vbRed
    ElseIf dif > 0 Then
        SetColor lblTag, vbGreen
    Else
        SetColor lblTag, vbBlack
    End If
End Sub


'------------------Загрузка и сохранение данных-----------------------------------
Public Function LoadTimes() As F_Times
    'Загружаем времена из фигуры
    LoadTime(fir).LoadTime(fin).LoadTime(inf).LoadTime(fArr).LoadTime(fStv).LoadTime(loc).LoadTime(log).LoadTime(lpp).LoadTime (endF)
Set LoadTimes = Me
End Function
Public Function LoadTime(ByVal tag As String) As F_Times
'Dim str As String
    'Загружаем время указанного показателя
'    str = CellVal(formShp, tag, visUnitsString)
    SetVal tag, CellVal(formShp, tag, visUnitsString)
Set LoadTime = Me
End Function

Public Function SaveTimes() As F_Times
    'Сохраняем времена
    SaveTime(fir).SaveTime(fin).SaveTime(inf).SaveTime(fArr).SaveTime(fStv).SaveTime(loc).SaveTime(log).SaveTime(lpp).SaveTime (endF)
Set SaveTimes = Me
End Function
Public Function SaveTime(ByVal tag As String) As F_Times
    'Сохраняем время указанного показателя
    SetCellVal formShp, tag, GetVal(tag)
Set SaveTime = Me
End Function

'------------------Манипуляции с данными-----------------------------------
Private Sub ClearDate(ByVal tag As String)
    GetControl(tag).Text = ""
End Sub

Private Sub TimeAdd(ByVal tag As String, Optional ByVal val As Integer = 1)
Dim dt As Date
Dim cntrl As Control
Dim str As String
    
    On Error GoTo EX
    
    Set cntrl = GetControl(tag)
    dt = CDate(cntrl.value)
    dt = DateAdd("n", val, dt)
    
    cntrl.Text = CStr(dt)
Exit Sub
EX:
    If Err.number = 13 Then
        str = GetControl(tag & ".cpt").Caption
        MsgBox "Не верно указана дата в поле '" & Left(str, Len(str) - 1) & "'"
    End If
End Sub

'------------------Анализ модели----------------------
Private Function GetMinArrTime() As Date
Dim dt As Date
Dim dtMin As Date
Dim shp As Visio.Shape
    
    For Each shp In Application.ActivePage.Shapes
        If CellVal(shp, "Actions.MainManeure.Checked") = 0 Then
            dt = CellVal(shp, "Prop.ArrivalTime", visDate)
            If dt > 0 Then
                If dtMin = 0 Or dtMin > dt Then dtMin = dt
            End If
        End If
    Next shp
    
GetMinArrTime = dtMin
End Function
Private Function GetMinStvTime() As Date
Dim dt As Date
Dim dtMin As Date
Dim shp As Visio.Shape
    
    For Each shp In Application.ActivePage.Shapes
        If CellVal(shp, "Actions.MainManeure.Checked") = 0 Then
            If IsGFSShapeWithIP(shp, Array(ipStvolRuch, ipStvolRuchPena, ipStvolLafVoda, ipStvolLafPena, ipStvolLafPoroshok, ipStvolLafVozimiy, ipStvolGas, ipStvolPoroshok)) Then
                dt = CellVal(shp, "Prop.SetTime", visDate)
                If dt > 0 Then
                    If dtMin = 0 Or dtMin > dt Then dtMin = dt
                End If
            End If
        End If
    Next shp
    
GetMinStvTime = dtMin
End Function

'------------------Проверка данных----------------------
Private Sub CheckField(ByVal tag As String)
Dim cntrl As Control
Dim str As String
    
    On Error GoTo EX
    
    Set cntrl = GetControl(tag)
    str = cntrl.Text
    
    If isCorrectData(str) Then
        SetColor tag, vbBlack
    Else
        SetColor tag, vbRed
    End If
    
Exit Sub
EX:
    Debug.Print "Error"
End Sub

''------------------Создание строк данных в странице----------------------
'Private Sub AddPageTimeProps()
'Dim tmpRowInd As Integer
'Dim tmpRow As Visio.Row
'Dim shp As Visio.Shape
'
''    On Error Resume Next
'
'    Set shp = Application.ActivePage.PageSheet
'
'    'Время возникновения
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время возникновения" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время возникновения пожара" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, fir, GetVal(fir)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время возникновения" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireTime" & Chr(34) & ",TheDoc!User.FireTime)+DEPENDSON(TheDoc!User.FireTime)"  'ВАЖНО !!! формулы вставлялись явным образом - текстом, иначе можно напороться на ошибки несовпадения индексов строк Scratch
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireTime" & Chr(34) & ", Prop.FireTime) + DEPENDSON(Prop.FireTime)"      'ВАЖНО !!! формулы вставлялись явным образом - текстом, иначе можно напороться на ошибки несовпадения индексов строк Scratch
'
'    'Время обнаружения
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "FindTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время обнаружения" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время обнаружения пожара" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, fin, GetVal(fin)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время обнаружения" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FindTime" & Chr(34) & ",TheDoc!User.FindTime)+DEPENDSON(TheDoc!User.FindTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FindTime" & Chr(34) & ", Prop.FindTime) + DEPENDSON(Prop.FindTime)"
'
'    'Время сообщения
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "InfoTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время сообщения" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время сообщения о пожаре" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, inf, GetVal(inf)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время сообщения" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.InfoTime" & Chr(34) & ",TheDoc!User.InfoTime)+DEPENDSON(TheDoc!User.InfoTime)"  '
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.InfoTime" & Chr(34) & ", Prop.InfoTime) + DEPENDSON(Prop.InfoTime)"
'
'    'Время прибытия первого подразделения
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstArrivalTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время прибытия" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время прибытия к месту пожара первого подразделения" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, fArr, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время прибытия" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstArrivalTime" & Chr(34) & ",TheDoc!User.FirstArrivalTime)+DEPENDSON(TheDoc!User.FirstArrivalTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstArrivalTime" & Chr(34) & ", Prop.FirstArrivalTime) + DEPENDSON(Prop.FirstArrivalTime)"
'
'    'Время подачи первого ствола
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstStvolTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время подачи ствола" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время подачи первого ствола" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, fStv, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время подачи ствола" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstStvolTime" & Chr(34) & ",TheDoc!User.FirstStvolTime)+DEPENDSON(TheDoc!User.FirstStvolTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstStvolTime" & Chr(34) & ", Prop.FirstStvolTime) + DEPENDSON(Prop.FirstStvolTime)"
'
'    'Время локализации
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "LocalizationTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время локализации" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время локализации" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, loc, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время локализации" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LocalizationTime" & Chr(34) & ",TheDoc!User.LocalizationTime)+DEPENDSON(TheDoc!User.LocalizationTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LocalizationTime" & Chr(34) & ", Prop.LocalizationTime) + DEPENDSON(Prop.LocalizationTime)"
'
'    'Время ЛОГ
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "LOGTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время ликвидации ОГ" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации открытого горения" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, log, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время ликвидации ОГ" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LOGTime" & Chr(34) & ",TheDoc!User.LOGTime)+DEPENDSON(TheDoc!User.LOGTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LOGTime" & Chr(34) & ", Prop.LOGTime) + DEPENDSON(Prop.LOGTime)"
'
'    'Время ЛПП
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "LPPTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время ликвидации ПП" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации последствий пожара" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, lpp, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время ликвидации ПП" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LPPTime" & Chr(34) & ",TheDoc!User.LPPTime)+DEPENDSON(TheDoc!User.LPPTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LPPTime" & Chr(34) & ", Prop.LPPTime) + DEPENDSON(Prop.LPPTime)"
'
'    'Время окончания пожара
'        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireEndTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время завершения работ" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации последствий пожара" & """"
'        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
'        SetCellVal shp, endF, GetVal(fArr)
'
'        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время завершения работ" & """"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireEndTime" & Chr(34) & ",TheDoc!User.FireEndTime)+DEPENDSON(TheDoc!User.FireEndTime)"
'        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireEndTime" & Chr(34) & ", Prop.FireEndTime) + DEPENDSON(Prop.FireEndTime)"
'
'End Sub
