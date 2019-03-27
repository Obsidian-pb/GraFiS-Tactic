Attribute VB_Name = "m_WorkWithCalculator"
Option Explicit

'----------------В модуле хранятся процедуры работы со калькулятором расчета параметров работы в СИЗОД--------------------------------------


Public Sub Ps_CalculatorShow(ShpObj As Visio.Shape, aS_SIZODType As String)
'Процедура показа формы калькулятора
Dim i As Integer
Dim vD_Time As Date

    On Error GoTo EX
'---Определяем стартовые значения формы
    ChainContentForm.VS_DevceType = aS_SIZODType
    ChainContentForm.VS_DeviceModel = ShpObj.Cells("Prop.AirDevice").ResultStr(visUnitsString)
    ChainContentForm.VL_InitShapeID = ShpObj.ID
    ChainContentForm.CB_Conditions = ShpObj.Cells("Prop.WorkPlace").ResultStr(visUnitsString)
    ChainContentForm.TB_DirectExpense = Int(ShpObj.Cells("Prop.AirConsuption").ResultStr(visNumber))
    ChainContentForm.ChkB_ShowResults = ShpObj.Cells("Actions.ResultShow.Checked").ResultStr(visNumber)
    ChainContentForm.VB_TimeChange = False
    ChainContentForm.VB_TimeArrivalChange = False
    vD_Time = ShpObj.Cells("Prop.FormingTime").ResultStr(visDate)
        ChainContentForm.TB_MainTimeEnter = _
        IIf(Hour(vD_Time) < 10, "0", "") & Hour(vD_Time) & ":" _
        & IIf(Minute(vD_Time) < 10, "0", "") & Minute(vD_Time) & ":" _
        & IIf(Second(vD_Time) < 10, "0", "") & Second(vD_Time)
    If ShpObj.CellExists("Prop.ArrivalTime", 0) = True Then 'Если поле имеется в текущей фигуре
        vD_Time = ShpObj.Cells("Prop.ArrivalTime").ResultStr(visDate)
        ChainContentForm.TB_TimeArrival = _
            IIf(Hour(vD_Time) < 10, "0", "") & Hour(vD_Time) & ":" _
            & IIf(Minute(vD_Time) < 10, "0", "") & Minute(vD_Time) & ":" _
            & IIf(Second(vD_Time) < 10, "0", "") & Second(vD_Time)
    Else 'Если отсутствует
        vD_Time = ShpObj.Cells("Prop.FormingTime").ResultStr(visDate)
        ChainContentForm.TB_TimeArrival = _
            IIf(Hour(vD_Time) < 10, "0", "") & Hour(vD_Time) & ":" _
            & IIf(Minute(vD_Time) < 10, "0", "") & Minute(vD_Time) & ":" _
            & IIf(Second(vD_Time) < 10, "0", "") & Second(vD_Time)
    End If
    
    '---Экспортируем данные для газодымозащитника №1
    ChainContentForm.TB_Perc1.Value = ShpObj.Cells("Scratch.A1").ResultStr(visUnitsString)
    ChainContentForm.TB_Perc1_P1.Value = Int(ShpObj.Cells("Scratch.B1").ResultStr(visNumber))
    ChainContentForm.TB_Perc1_P2.Value = Int(ShpObj.Cells("Scratch.C1").ResultStr(visNumber))
    '---Экспортируем данные для газодымозащитника №2
    If Not (ShpObj.Cells("Scratch.B2").ResultStr(visUnitsString) = "" Or ShpObj.Cells("Scratch.C2").ResultStr(visUnitsString) = "") Then
        ChainContentForm.TB_Perc2.Value = ShpObj.Cells("Scratch.A2").ResultStr(visUnitsString)
        ChainContentForm.TB_Perc2_P1.Value = Int(ShpObj.Cells("Scratch.B2").ResultStr(visNumber))
        ChainContentForm.TB_Perc2_P2.Value = Int(ShpObj.Cells("Scratch.C2").ResultStr(visNumber))
    End If
    '---Экспортируем данные для газодымозащитника №3
    If Not (ShpObj.Cells("Scratch.B3").ResultStr(visUnitsString) = "" Or ShpObj.Cells("Scratch.C3").ResultStr(visUnitsString) = "") Then
        ChainContentForm.TB_Perc3.Value = ShpObj.Cells("Scratch.A3").ResultStr(visUnitsString)
        ChainContentForm.TB_Perc3_P1.Value = Int(ShpObj.Cells("Scratch.B3").ResultStr(visNumber))
        ChainContentForm.TB_Perc3_P2.Value = Int(ShpObj.Cells("Scratch.C3").ResultStr(visNumber))
        ChainContentForm.ChkB_Perc3 = True
    Else
        ChainContentForm.TB_Perc3.Value = ShpObj.Cells("Scratch.A3").ResultStr(visUnitsString)
        ChainContentForm.ChkB_Perc3 = False
    End If
    
    '---Экспортируем данные для газодымозащитника №4
    If Not (ShpObj.Cells("Scratch.B4").ResultStr(visUnitsString) = "" Or ShpObj.Cells("Scratch.C4").ResultStr(visUnitsString) = "") Then
        ChainContentForm.TB_Perc4.Value = ShpObj.Cells("Scratch.A4").ResultStr(visUnitsString)
        ChainContentForm.TB_Perc4_P1.Value = Int(ShpObj.Cells("Scratch.B4").ResultStr(visNumber))
        ChainContentForm.TB_Perc4_P2.Value = Int(ShpObj.Cells("Scratch.C4").ResultStr(visNumber))
        ChainContentForm.ChkB_Perc4 = True
    Else
        ChainContentForm.TB_Perc4.Value = ShpObj.Cells("Scratch.A4").ResultStr(visUnitsString)
        ChainContentForm.ChkB_Perc4 = False
    End If
    '---Экспортируем данные для газодымозащитника №5
    If Not (ShpObj.Cells("Scratch.B5").ResultStr(visUnitsString) = "" Or ShpObj.Cells("Scratch.C5").ResultStr(visUnitsString) = "") Then
        ChainContentForm.TB_Perc5.Value = ShpObj.Cells("Scratch.A5").ResultStr(visUnitsString)
        ChainContentForm.TB_Perc5_P1.Value = Int(ShpObj.Cells("Scratch.B5").ResultStr(visNumber))
        ChainContentForm.TB_Perc5_P2.Value = Int(ShpObj.Cells("Scratch.C5").ResultStr(visNumber))
        ChainContentForm.ChkB_Perc5 = True
    Else
        ChainContentForm.TB_Perc5.Value = ShpObj.Cells("Scratch.A5").ResultStr(visUnitsString)
        ChainContentForm.ChkB_Perc5 = False
    End If
    '---Экспортируем данные для газодымозащитника №6
    If Not (ShpObj.Cells("Scratch.B6").ResultStr(visUnitsString) = "" Or ShpObj.Cells("Scratch.C6").ResultStr(visUnitsString) = "") Then
        ChainContentForm.TB_Perc6.Value = ShpObj.Cells("Scratch.A6").ResultStr(visUnitsString)
        ChainContentForm.TB_Perc6_P1.Value = Int(ShpObj.Cells("Scratch.B6").ResultStr(visNumber))
        ChainContentForm.TB_Perc6_P2.Value = Int(ShpObj.Cells("Scratch.C6").ResultStr(visNumber))
        ChainContentForm.ChkB_Perc6 = True
    Else
        ChainContentForm.TB_Perc6.Value = ShpObj.Cells("Scratch.A6").ResultStr(visUnitsString)
        ChainContentForm.ChkB_Perc6 = False
    End If

''---Показываем форму
    ChainContentForm.Show

Exit Sub
EX:
    SaveLog Err, "Ps_CalculatorShow", ShpObj.Name
End Sub





