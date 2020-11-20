VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Times 
   Caption         =   "Временные показатели модели"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8505
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
            MsgBox "На странице нет фигур со свойством 'Время прибытия'."
        Else
            SetVal fStv, CStr(dt)
        End If
    End Sub
    '-------------Определение разниц между временами--------------
    Private Sub txt_FindTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_FireEndTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_FireTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_FirstArrivalTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_FirstStvolTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_InfoTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_LocalizationTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_LOGTime_Change()
        CheckDiffs
    End Sub
    Private Sub txt_LPPTime_Change()
        CheckDiffs
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
    On Error GoTo ex
    test = CDate(txt)
    
isCorrectData = True
Exit Function
ex:
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
    
    On Error GoTo ex
    
    Set cntrl = GetControl(tag)
    dt = CDate(cntrl.value)
    dt = DateAdd("n", val, dt)
    
    cntrl.Text = CStr(dt)
Exit Sub
ex:
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
        dt = CellVal(shp, "Prop.ArrivalTime", visDate)
        If dt > 0 Then
            If dtMin = 0 Or dtMin > dt Then dtMin = dt
        End If
    Next shp
    
GetMinArrTime = dtMin
End Function
Private Function GetMinStvTime() As Date
Dim dt As Date
Dim dtMin As Date
Dim shp As Visio.Shape
    
    For Each shp In Application.ActivePage.Shapes
        If IsGFSShapeWithIP(shp, Array(ipStvolRuch, ipStvolRuchPena, ipStvolLafVoda, ipStvolLafPena, ipStvolLafPoroshok, ipStvolLafVozimiy, ipStvolGas, ipStvolPoroshok)) Then
            dt = CellVal(shp, "Prop.SetTime", visDate)
            If dt > 0 Then
                If dtMin = 0 Or dtMin > dt Then dtMin = dt
            End If
        End If
    Next shp
    
GetMinStvTime = dtMin
End Function





