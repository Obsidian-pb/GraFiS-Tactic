VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChainContentForm 
   Caption         =   "Калькулятор показателей работы в СИЗОД"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   OleObjectBlob   =   "ChainContentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChainContentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VS_DevceType As String
Public VS_DeviceModel As String
Public VL_InitShapeID As Long
Public VB_TimeChange As Boolean
Public VB_TimeArrivalChange As Boolean






'------------------------------------------Базовые процедуры формы----------------------------------------------
Private Sub CB_Conditions_Change()
    ps_ResultsRecalc
End Sub



Private Sub TB_DirectExpense_Change()
    On Error GoTo ex
    TB_DirectExpense.Value = CInt(TB_DirectExpense.Value)
'    If Not (TB_DirectExpense = "" Or TB_DirectExpense = 0) Then ps_ResultsRecalc
    ps_ResultsRecalc
ex:
    Exit Sub
End Sub

Private Sub TB_MainTimeEnter_Change()
Dim vDt_Time As Date
    On Error GoTo ex
    Me.VB_TimeChange = True
    vDt_Time = _
        DateAdd("s", CDbl(TB_TimeTotal.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
    Me.TB_MainTimeExit.Value = _
        IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
        & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
        & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
ex:
    Exit Sub
End Sub

Private Sub TB_TimeArrival_Change()
'Процедура события изменения времени прибытия к очагу
Dim vDt_Time As Date
    On Error GoTo ex
    Me.VB_TimeArrivalChange = True
    '---Вычисляем время подачи команды в случае, если очаг обнаружен
        vDt_Time = _
            DateAdd("s", CDbl(Me.TB_TimeAtFire.Value) * 60, TimeValue(Me.TB_TimeArrival))
        Me.TB_OrderTimeAtFire.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
ex:
    Exit Sub

End Sub

Private Sub UserForm_Activate()
    
    ps_Conditions_ListFill
    ps_Models_ListFill

    ps_ResultsRecalc
    
End Sub

Private Sub CB_Quit_Click()
    Me.Hide
End Sub

Private Sub CB_OK_Click()
    ps_ValuesBack
    Me.Hide
End Sub

'------------------------------------------Процедуры настройки формы----------------------------------------------
Private Sub ps_Conditions_ListFill()
'Процедура заполнения выпадающего списка условий работы
    With CB_Conditions
        .Clear
        .AddItem "Стандартные условия"
        .AddItem "Сложные условия"
        .ListIndex = 0
    End With

End Sub

Private Sub ps_Models_ListFill()
'Процедура получения данных по аппаратам и заполнения на их основе списка
Dim dbs As Object
Dim rst As Object
Dim vS_Path As String
Dim i As Integer
    
    On Error GoTo ex
    
'---Очищаем список от прежних значений
LB_DeviceModel.Clear

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных в зависимости от типа аппаратов
        If VS_DevceType = "ДАСВ" Then
            SQLQuery = "SELECT ДАСВ.Модель, ДАСВ.[Объем баллонов], ДАСВ.[Давление редуктора], Баллоны.Ксж " _
            & "FROM Баллоны RIGHT JOIN ДАСВ ON Баллоны.КодБаллона = ДАСВ.Баллон ORDER BY ДАСВ.Модель;"
        Else
            SQLQuery = "SELECT ДАСК.Модель, ДАСК.[Объем баллонов], ДАСК.[Давление редуктора], ДАСК.Ксж " _
            & "FROM ДАСК;"
        End If
        
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
        i = 0
        Do Until .EOF
            LB_DeviceModel.AddItem
            LB_DeviceModel.Column(0, i) = !Модель
            LB_DeviceModel.Column(1, i) = CStr(![Объем баллонов])
            LB_DeviceModel.Column(2, i) = ![Давление редуктора]
            LB_DeviceModel.Column(3, i) = CStr(!Ксж)
            i = i + 1
            .MoveNext
        Loop
    End With
    
    For i = 0 To Me.LB_DeviceModel.ListCount - 1
        If Me.LB_DeviceModel.Column(0, i) = VS_DeviceModel Then
            Me.LB_DeviceModel.Value = i
            Exit For
        End If
    Next i
    
'---Очищаем объектные переменные
    pth = ""
    Set dbs = Nothing
    Set rst = Nothing
   
Exit Sub
ex:
    SaveLog Err, "ps_Models_ListFill"
End Sub

Private Sub LB_DeviceModel_Change()
    If LB_DeviceModel.ListCount > 0 Then
        TB_BallonsValue = LB_DeviceModel.Column(1, LB_DeviceModel.Value)
        TB_ReductorNeedPressure = LB_DeviceModel.Column(2, LB_DeviceModel.Value)
        TB_CompFactor = LB_DeviceModel.Column(3, LB_DeviceModel.Value)
        '---Обновляем поля результатов расчета
        ps_ResultsRecalc
    End If
End Sub



Private Sub ChkB_Perc3_Change()
    TB_Perc3.Enabled = ChkB_Perc3.Value
    TB_Perc3_P1.Enabled = ChkB_Perc3.Value
    TB_Perc3_P2.Enabled = ChkB_Perc3.Value
    
    If ChkB_Perc3.Value = True Then
        TB_Perc3.BackStyle = fmBackStyleOpaque
        TB_Perc3_P1.BackStyle = fmBackStyleOpaque
        TB_Perc3_P2.BackStyle = fmBackStyleOpaque
        TB_Perc3_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc3.BackStyle = fmBackStyleTransparent
        TB_Perc3_P1.BackStyle = fmBackStyleTransparent
        TB_Perc3_P2.BackStyle = fmBackStyleTransparent
        TB_Perc3_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc4_Change()
    TB_Perc4.Enabled = ChkB_Perc4.Value
    TB_Perc4_P1.Enabled = ChkB_Perc4.Value
    TB_Perc4_P2.Enabled = ChkB_Perc4.Value

    If ChkB_Perc4.Value = True Then
        TB_Perc4.BackStyle = fmBackStyleOpaque
        TB_Perc4_P1.BackStyle = fmBackStyleOpaque
        TB_Perc4_P2.BackStyle = fmBackStyleOpaque
        TB_Perc4_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc4.BackStyle = fmBackStyleTransparent
        TB_Perc4_P1.BackStyle = fmBackStyleTransparent
        TB_Perc4_P2.BackStyle = fmBackStyleTransparent
        TB_Perc4_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc5_Change()
    TB_Perc5.Enabled = ChkB_Perc5.Value
    TB_Perc5_P1.Enabled = ChkB_Perc5.Value
    TB_Perc5_P2.Enabled = ChkB_Perc5.Value

    If ChkB_Perc5.Value = True Then
        TB_Perc5.BackStyle = fmBackStyleOpaque
        TB_Perc5_P1.BackStyle = fmBackStyleOpaque
        TB_Perc5_P2.BackStyle = fmBackStyleOpaque
        TB_Perc5_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc5.BackStyle = fmBackStyleTransparent
        TB_Perc5_P1.BackStyle = fmBackStyleTransparent
        TB_Perc5_P2.BackStyle = fmBackStyleTransparent
        TB_Perc5_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc6_Change()
    TB_Perc6.Enabled = ChkB_Perc6.Value
    TB_Perc6_P1.Enabled = ChkB_Perc6.Value
    TB_Perc6_P2.Enabled = ChkB_Perc6.Value

    If ChkB_Perc6.Value = True Then
        TB_Perc6.BackStyle = fmBackStyleOpaque
        TB_Perc6_P1.BackStyle = fmBackStyleOpaque
        TB_Perc6_P2.BackStyle = fmBackStyleOpaque
        TB_Perc6_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc6.BackStyle = fmBackStyleTransparent
        TB_Perc6_P1.BackStyle = fmBackStyleTransparent
        TB_Perc6_P2.BackStyle = fmBackStyleTransparent
        TB_Perc6_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

'------------------------------------------Процедуры вычисления показателей давления--------------------------------
Private Sub TB_Perc1_P1_Change()
    If Not TB_Perc1_P1 = "" Then TB_Perc1_PFall = TB_Perc1_P1 - TB_Perc1_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc1_P2_Change()
    If Not TB_Perc1_P2 = "" Then TB_Perc1_PFall = TB_Perc1_P1 - TB_Perc1_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc2_P1_Change()
    If Not TB_Perc2_P1 = "" Then TB_Perc2_PFall = TB_Perc2_P1 - TB_Perc2_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc2_P2_Change()
    If Not TB_Perc2_P2 = "" Then TB_Perc2_PFall = TB_Perc2_P1 - TB_Perc2_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc3_P1_Change()
    If Not TB_Perc3_P1 = "" Then TB_Perc3_PFall = TB_Perc3_P1 - TB_Perc3_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc3_P2_Change()
    If Not TB_Perc3_P2 = "" Then TB_Perc3_PFall = TB_Perc3_P1 - TB_Perc3_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc4_P1_Change()
    If Not TB_Perc4_P1 = "" Then TB_Perc4_PFall = TB_Perc4_P1 - TB_Perc4_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc4_P2_Change()
    If Not TB_Perc4_P2 = "" Then TB_Perc4_PFall = TB_Perc4_P1 - TB_Perc4_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc5_P1_Change()
    If Not TB_Perc5_P1 = "" Then TB_Perc5_PFall = TB_Perc5_P1 - TB_Perc5_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc5_P2_Change()
    If Not TB_Perc5_P2 = "" Then TB_Perc5_PFall = TB_Perc5_P1 - TB_Perc5_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc6_P1_Change()
    If Not TB_Perc6_P1 = "" Then TB_Perc6_PFall = TB_Perc6_P1 - TB_Perc6_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc6_P2_Change()
    If Not TB_Perc6_P2 = "" Then TB_Perc6_PFall = TB_Perc6_P1 - TB_Perc6_P2
    ps_ResultsRecalc
End Sub

Private Sub Min_P1()
'Процедура вычисления минимального давления при включении
Dim x(6) As Integer
Dim i As Integer
Dim min As Integer

    If Not (TB_Perc1_P1 = "" Or TB_Perc1_P1 = 0) Then x(0) = TB_Perc1_P1 Else x(0) = 300
    If Not (TB_Perc2_P1 = "" Or TB_Perc2_P1 = 0) Then x(1) = TB_Perc2_P1 Else x(1) = 300
    If Not (TB_Perc3_P1 = "" Or TB_Perc3_P1 = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_P1 Else x(2) = 300
    If Not (TB_Perc4_P1 = "" Or TB_Perc4_P1 = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_P1 Else x(3) = 300
    If Not (TB_Perc5_P1 = "" Or TB_Perc5_P1 = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_P1 Else x(4) = 300
    If Not (TB_Perc6_P1 = "" Or TB_Perc6_P1 = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_P1 Else x(5) = 300
    
    min = x(0)
    For i = 0 To 5
        If x(i) < min Then min = x(i)
    Next i
    
    TB_P1_Min = min
End Sub

Private Sub Min_P2()
'Процедура вычисления минимального давления у очага пожара
Dim x(6) As Integer
Dim i As Integer
Dim min As Integer

    If Not (TB_Perc1_P2 = "" Or TB_Perc1_P2 = 0) Then x(0) = TB_Perc1_P2 Else x(0) = 300
    If Not (TB_Perc2_P2 = "" Or TB_Perc2_P2 = 0) Then x(1) = TB_Perc2_P2 Else x(1) = 300
    If Not (TB_Perc3_P2 = "" Or TB_Perc3_P2 = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_P2 Else x(2) = 300
    If Not (TB_Perc4_P2 = "" Or TB_Perc4_P2 = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_P2 Else x(3) = 300
    If Not (TB_Perc5_P2 = "" Or TB_Perc5_P2 = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_P2 Else x(4) = 300
    If Not (TB_Perc6_P2 = "" Or TB_Perc6_P2 = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_P2 Else x(5) = 300
    
    min = x(0)
    For i = 0 To 5
        If x(i) < min Then min = x(i)
    Next i
    
    TB_P2_Min = min
End Sub

Private Sub Max_PFall()
'Процедура вычисления максимального падения давления на пути к очагу пожара
Dim x(6) As Integer
Dim i As Integer
Dim max As Integer

    If Not (TB_Perc1_PFall = "" Or TB_Perc1_PFall = 0) Then x(0) = TB_Perc1_PFall Else x(0) = 0
    If Not (TB_Perc2_PFall = "" Or TB_Perc2_PFall = 0) Then x(1) = TB_Perc2_PFall Else x(1) = 0
    If Not (TB_Perc3_PFall = "" Or TB_Perc3_PFall = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_PFall Else x(2) = 0
    If Not (TB_Perc4_PFall = "" Or TB_Perc4_PFall = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_PFall Else x(3) = 0
    If Not (TB_Perc5_PFall = "" Or TB_Perc5_PFall = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_PFall Else x(4) = 0
    If Not (TB_Perc6_PFall = "" Or TB_Perc6_PFall = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_PFall Else x(5) = 0
    
    max = x(0)
    For i = 0 To 5
        If x(i) > max Then max = x(i)
    Next i
    
    TB_PFall_Max = max
End Sub

'------------------------------------------Процедуры проведения собственно расчетов--------------------------------
Private Function TotalWorkTimeCalculate(ai_minPressure As Integer, ai_ReductorPressure As Integer, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'Процедура вычисления общего времени работы (в секундах)
Dim vd_Temp As Double
    'Вычисляем промежуточное значение времени работы в секундах
    vd_Temp = (ai_minPressure - ai_ReductorPressure) * as_BalloonValue / ((ai_AirExpence / 60) * as_CempressFaxtor)
    'Возвращаем полученное значение
    TotalWorkTimeCalculate = CLng(vd_Temp)
End Function

Private Function TimeAtFireCalculate(ai_minPressure As Integer, as_BackResrve As Single, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'Процедура вычисления контрольного даления выхода (в секундах)
Dim vd_Temp As Double
    'Вычисляем промежуточное значение времени работы в секундах
    vd_Temp = (ai_minPressure - as_BackResrve) * as_BalloonValue / ((ai_AirExpence / 60) * as_CempressFaxtor)
    'Возвращаем полученное значение
    TimeAtFireCalculate = CLng(vd_Temp)
End Function

Private Function fs_WorkTimeCalculateUntilFireFind(ai_maxFallPressure As Integer, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'Процедура вычисления времени работы с момента включения до момента подачи команды постовым при необнаружении очага (в секундах)
Dim vd_Temp As Double
    'Вычисляем промежуточное значение времени работы в секундах
    vd_Temp = (ai_maxFallPressure * as_BalloonValue) / ((ai_AirExpence / 60) * as_CempressFaxtor)
    'Возвращаем полученное значение
    fs_WorkTimeCalculateUntilFireFind = CLng(vd_Temp)
End Function

Private Function fs_BackResrveCalculate(ai_maxFallPressure As Integer, ab_HardFactor As Boolean, _
                                        ai_ReductorPressure As Integer) As Single
'Процедура вычисления запаса давления необходимого для возвращения, атм
    If ab_HardFactor = True Then
        fs_BackResrveCalculate = ai_maxFallPressure * 2 + ai_ReductorPressure
    Else
        fs_BackResrveCalculate = ai_maxFallPressure * 1.5 + ai_ReductorPressure
    End If
End Function

Private Function fs_MaxPressureFallWF(ai_minEnterPressure As Integer, ab_HardFactor As Boolean, _
                                        ai_ReductorPressure As Integer) As Single
'Процедура вычисления максимально возможного падения давления в аппаратах, атм
    If ab_HardFactor = True Then
        fs_MaxPressureFallWF = (ai_minEnterPressure - ai_ReductorPressure) / 3
    Else
        fs_MaxPressureFallWF = (ai_minEnterPressure - ai_ReductorPressure) / 2.5
    End If
End Function

Private Function fB_ConvHardFactor(aS_HardFactor As String) As Boolean
'Функция возвращает ИСТИНА, если сложные условия и ЛОЖЬ, если стандартные
    If aS_HardFactor = "Стандартные условия" Then
        fB_ConvHardFactor = False
    Else
        fB_ConvHardFactor = True
    End If
End Function

Private Sub ps_ResultsRecalc()
'Процедура пересчета результатов расчета
Dim vDt_Time As Date
    
    Min_P1
    Min_P2
    Max_PFall
    
    On Error GoTo ex
    
    '---Вычисляем время работы при НЕ обнаруженном очаге
    TB_MaxFall = _
        CStr(fs_MaxPressureFallWF(CInt(TB_P1_Min), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))

    TB_ExitPressure = _
        CStr(CInt(TB_P1_Min) - fs_MaxPressureFallWF(CInt(TB_P1_Min), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))
        
    TB_TimuUntilOrder = _
        CStr(Round(fs_WorkTimeCalculateUntilFireFind(CInt(TB_MaxFall), _
        CSng(TB_BallonsValue), CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    
    '---Вычисляем время работы при обнаруженном очаге
    TB_BackWayReserv = _
        CStr(fs_BackResrveCalculate(CInt(TB_PFall_Max), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))
    TB_TimeAtFire.Value = _
        CStr(Round(TimeAtFireCalculate(CInt(TB_P2_Min), CSng(TB_BackWayReserv), CSng(TB_BallonsValue), _
        CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    TB_TimeTotal.Value = _
        CStr(Round(TotalWorkTimeCalculate(CInt(TB_P1_Min), CInt(TB_ReductorNeedPressure), CSng(TB_BallonsValue), _
        CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    On Error Resume Next
    '---Вычисляем время подачи команды в случае, если очаг НЕ будет обнаружен
        vDt_Time = _
            DateAdd("s", CDbl(TB_TimuUntilOrder.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
        Me.TB_OrderTime.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
    '---Вычисляем время подачи команды в случае, если очаг обнаружен
        vDt_Time = _
            DateAdd("s", CDbl(Me.TB_TimeAtFire.Value) * 60, TimeValue(Me.TB_TimeArrival))
        Me.TB_OrderTimeAtFire.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
    '---Вычисляем время выхода
        vDt_Time = _
            DateAdd("s", CDbl(TB_TimeTotal.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
        Me.TB_MainTimeExit.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)

Exit Sub
ex:
    SaveLog Err, "ps_ResultsRecalc"
End Sub

'-------------------------------Процедуры возвращения данных в инициирововшую фигуру--------------------------------
Private Sub ps_ValuesBack()
'Процедура возвращения данных в фигуру-инициатор
Dim vO_IntShape As Visio.Shape
Dim vD_TimeTemp As Date
Dim vL_Interval As Long
Dim vdbl_TimeTemp As Double

    On Error GoTo ex

Set vO_IntShape = Application.ActivePage.Shapes.ItemFromID(VL_InitShapeID)

'---Возвращаем значения свойств фигур .Prop
'    ChainContentForm.VS_DevceType = aS_SIZODType
    vO_IntShape.Cells("Prop.AirDevice").FormulaU = """" & Me.LB_DeviceModel.Column(0, Me.LB_DeviceModel.Value) & """"
'    ChainContentForm.VL_InitShapeID = ShpObj.ID
    vO_IntShape.Cells("Prop.Personnel").FormulaU = 2 + IIf(Me.ChkB_Perc3.Value, 1, 0) + IIf(Me.ChkB_Perc4.Value, 1, 0) + _
        IIf(Me.ChkB_Perc5.Value, 1, 0) + IIf(Me.ChkB_Perc6.Value, 1, 0)
    vO_IntShape.Cells("Prop.WorkPlace").FormulaU = """" & Me.CB_Conditions & """"
    vO_IntShape.Cells("Prop.AirConsuption").FormulaU = CStr(Me.TB_DirectExpense)
    vO_IntShape.Cells("Actions.ResultShow.Checked").FormulaU = Me.ChkB_ShowResults.Value
    If Me.VB_TimeChange = True Then 'Обновляем данные по времени включения, в случае, если они менялись
        vD_TimeTemp = DateValue(vO_IntShape.Cells("Prop.FormingTime").ResultStr(visDate))
        vL_Interval = Hour(Me.TB_MainTimeEnter.Value) * 3600 + _
            Minute(Me.TB_MainTimeEnter.Value) * 60 + Second(Me.TB_MainTimeEnter.Value)
        vdbl_TimeTemp = CDbl(DateAdd("s", vL_Interval, vD_TimeTemp))
        vO_IntShape.Cells("Prop.FormingTime").FormulaU = "DATETIME(" & str(vdbl_TimeTemp) & ")"
    End If
    If Me.VB_TimeArrivalChange = True And vO_IntShape.CellExists("Prop.ArrivalTime", 0) = True Then 'Обновляем данные по времени прибытия к очагу, в случае, если они менялись и имеется соответствующая ячейка
        vD_TimeTemp = DateValue(vO_IntShape.Cells("Prop.ArrivalTime").ResultStr(visDate))
        vL_Interval = Hour(Me.TB_TimeArrival.Value) * 3600 + _
            Minute(Me.TB_TimeArrival.Value) * 60 + Second(Me.TB_TimeArrival.Value)
        vdbl_TimeTemp = CDbl(DateAdd("s", vL_Interval, vD_TimeTemp))
        vO_IntShape.Cells("Prop.ArrivalTime").FormulaU = "DATETIME(" & str(vdbl_TimeTemp) & ")"
    End If

    '---Возвращаем данные для газодымозащитника №1
    vO_IntShape.Cells("Scratch.A1").FormulaU = """" & Me.TB_Perc1.Value & """"
    vO_IntShape.Cells("Scratch.B1").FormulaU = Me.TB_Perc1_P1.Value
    vO_IntShape.Cells("Scratch.C1").FormulaU = Me.TB_Perc1_P2.Value
    '---Экспортируем данные для газодымозащитника №2
    If Me.ChkB_Perc2.Value Then
        vO_IntShape.Cells("Scratch.A2").FormulaU = """" & Me.TB_Perc2.Value & """"
        vO_IntShape.Cells("Scratch.B2").FormulaU = Me.TB_Perc2_P1.Value
        vO_IntShape.Cells("Scratch.C2").FormulaU = Me.TB_Perc2_P2.Value
    End If
    '---Экспортируем данные для газодымозащитника №3
    If Me.ChkB_Perc3.Value Then
        vO_IntShape.Cells("Scratch.A3").FormulaU = """" & Me.TB_Perc3.Value & """"
        vO_IntShape.Cells("Scratch.B3").FormulaU = Me.TB_Perc3_P1.Value
        vO_IntShape.Cells("Scratch.C3").FormulaU = Me.TB_Perc3_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B3").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C3").FormulaU = """" & """"
    End If

    '---Экспортируем данные для газодымозащитника №4
    If Me.ChkB_Perc4.Value Then
        vO_IntShape.Cells("Scratch.A4").FormulaU = """" & Me.TB_Perc4.Value & """"
        vO_IntShape.Cells("Scratch.B4").FormulaU = Me.TB_Perc4_P1.Value
        vO_IntShape.Cells("Scratch.C4").FormulaU = Me.TB_Perc4_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B4").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C4").FormulaU = """" & """"
    End If
    '---Экспортируем данные для газодымозащитника №5
    If Me.ChkB_Perc5.Value Then
        vO_IntShape.Cells("Scratch.A5").FormulaU = """" & Me.TB_Perc5.Value & """"
        vO_IntShape.Cells("Scratch.B5").FormulaU = Me.TB_Perc5_P1.Value
        vO_IntShape.Cells("Scratch.C5").FormulaU = Me.TB_Perc5_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B5").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C5").FormulaU = """" & """"
    End If
    '---Экспортируем данные для газодымозащитника №6
    If Me.ChkB_Perc6.Value Then
        vO_IntShape.Cells("Scratch.A6").FormulaU = """" & Me.TB_Perc6.Value & """"
        vO_IntShape.Cells("Scratch.B6").FormulaU = Me.TB_Perc6_P1.Value
        vO_IntShape.Cells("Scratch.C6").FormulaU = Me.TB_Perc6_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B6").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C6").FormulaU = """" & """"
    End If

Exit Sub
ex:
    SaveLog Err, "ps_ValuesBack"
End Sub





