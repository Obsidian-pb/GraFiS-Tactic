Attribute VB_Name = "m_NRS"
Option Explicit

'-------------------Модуль для хранения процедур и функций по анализу насосно-рукавных систем--------------------
Private shapesInNRS As Collection

'---Постоянные индеков графиков
'Const ccs_InIdent = "Connections.GFS_In"
'Const ccs_OutIdent = "Connections.GFS_Ou"
Const vb_ShapeType_Other = 0    'Ничего
Const vb_ShapeType_Hose = 1     'Рукава
Const vb_ShapeType_PTV = 2      'ПТВ
Const vb_ShapeType_Razv = 3     'Разветвление
Const vb_ShapeType_Tech = 4     'Техника
Const vb_ShapeType_VsasSet = 5  'Всасывающая сетка с линией
Const vb_ShapeType_GE = 6       'Гидроэлеватор

Const CP_GrafisVersion = 1      'Версия набора

'-----------------------общие сведения анализа-----------------------
Public PA_Count As Integer
Public MP_Count As Integer
Public Ship_Count As Integer
Public Train_Count As Integer

Public Hose38_Count As Integer
Public Hose51_Count As Integer
Public Hose66_Count As Integer
Public Hose77_Count As Integer
Public Hose89_Count As Integer
Public Hose110_Count As Integer
Public Hose150_Count As Integer
Public Hose200_Count As Integer
Public Hose250_Count As Integer
Public Hose300_Count As Integer
Public OtherHoses_Count As Integer

Public NapHoses_Lenight As Integer
Public VsasHoses_Lenight As Integer

Public Hose77NV_Count As Integer
Public Hose125NV_Count As Integer
Public Hose150NV_Count As Integer
Public Hose200NV_Count As Integer

Public Razv_Count As Integer
Public VS_Count As Integer
Public Collector As Integer
Public GE_Count As Integer
Public PS_Count As Integer
Public VsasSetc_Count As Integer
Public Kol_Count As Integer
Public PV_Count As Integer

Public StvA_Count As Integer
Public StvB_Count As Integer
Public StvLaf_Count As Integer
Public StvPen_Count As Integer
Public StvGPS_Count As Integer

Public PodOut As Double
Public PodIn As Double
Public HosesValue As Double
Public WaterValue As Double 'Объем воды в емкостях

Public PG_Count As Integer
Public PW_Count As Integer
Public PK_Count As Integer
Public WaterContainers_Count As Integer
Public WaterContainers_Value As Double
Public WaterOpen_Value As Double

Private PW_Value As Double




Public Sub GESystemTest(ShpObj As Visio.Shape)
'Основная процедура получения сведений о насосно-рукавной системе
    
    On Error GoTo EX
    
    '---Выделяем память под коллекцию фигур в НРС
    Set shapesInNRS = New Collection
    
    '---Наполняем коллекцию фигурами входящими в НРС
        GetTechShapeForGESystem ShpObj
    
    '---Анализируем фигуры в коллекции
        NRS_Analize
    '---Формируем отчет
        CreateReport
        
    Set shapesInNRS = Nothing
Exit Sub
EX:
    Set shapesInNRS = Nothing
End Sub

Private Sub GetTechShapeForGESystem(ByRef shp As Visio.Shape)
'Заполняем коллекцию фигур соединенных в НРС
Dim Con As Connect
Dim sideShp As Visio.Shape

    For Each Con In shp.Connects
        If Not IsShapeAllreadyChecked(Con.ToSheet) Then
            shapesInNRS.Add Con.ToSheet
            GetTechShapeForGESystem Con.ToSheet
        End If
    Next Con
    For Each Con In shp.FromConnects
        If Not IsShapeAllreadyChecked(Con.FromSheet) Then
            shapesInNRS.Add Con.FromSheet
            GetTechShapeForGESystem Con.FromSheet
        End If
    Next Con

End Sub

Private Function IsShapeAllreadyChecked(ByRef shp As Visio.Shape) As Boolean
'Функция возвращает Истину, если фигура уже имеется и ложь, если нет
Dim colShape As Visio.Shape

    For Each colShape In shapesInNRS
        If colShape = shp Then
            IsShapeAllreadyChecked = True
            Exit Function
        End If
    Next colShape
    
IsShapeAllreadyChecked = False
End Function

Private Sub CreateReport()
'Пркоа формирует и выводит отчет поНРС
Dim totalStr As String
Dim FlowOut As Double 'Скорость убывания жидкости
Dim DischargeTime As Double
    
    If PodOut > 0 Then totalStr = totalStr & "Общий расход системы - " & PodOut & " л/с" & Chr(10)
    If PodIn > 0 Then totalStr = totalStr & "Общий забор воды - " & PodIn & " л/с" & Chr(10)
    If HosesValue > 0 Then totalStr = totalStr & "Объем воды в рукавах - " & HosesValue & " л" & Chr(10)
    If WaterValue > 0 Then totalStr = totalStr & "Объем воды в емкостях МСП - " & WaterValue & " л" & Chr(10)
    If PW_Value > 0 Then totalStr = totalStr & "Объем доступной воды в пожарных водоемах - " & PW_Value & " л" & Chr(10)
    If WaterOpen_Value > 0 Then totalStr = totalStr & "Объем доступной воды в открытых водоисточниках - " & WaterOpen_Value & " л" & Chr(10)

    If PodOut > PodIn Then
        FlowOut = PodOut - PodIn
        DischargeTime = ((WaterValue + PW_Value + WaterOpen_Value - HosesValue) / FlowOut) / 60
        If Int(DischargeTime) > 0 Then
            totalStr = totalStr & "Возможное время работы системы - " & _
                     Int(DischargeTime) & ":" & Int((DischargeTime - Int(DischargeTime)) * 60) _
                     & Chr(10)
        Else
            totalStr = totalStr & "ОШИБКА РАСЧЕТА ВРЕМЕНИ РАБОТЫ СИСТЕМЫ!" & Chr(10)
        End If
    Else
        totalStr = totalStr & "Возможное время работы системы - бесконечно" & Chr(10)
    End If
    
    If PA_Count > 0 Then totalStr = totalStr & "Пожарных автомобилей - " & PA_Count & Chr(10)
    If MP_Count > 0 Then totalStr = totalStr & "Пожарных мотопомп - " & MP_Count & Chr(10)
    If Train_Count > 0 Then totalStr = totalStr & "Пожарных поездов - " & Train_Count & Chr(10)
    If Ship_Count > 0 Then totalStr = totalStr & "Пожарных судов - " & Ship_Count & Chr(10)
    
    If Hose38_Count > 0 Then totalStr = totalStr & "Напорные рукава 38 мм - " & Hose38_Count & Chr(10)
    If Hose51_Count > 0 Then totalStr = totalStr & "Напорные рукава 51 мм - " & Hose51_Count & Chr(10)
    If Hose66_Count > 0 Then totalStr = totalStr & "Напорные рукава 66 мм - " & Hose66_Count & Chr(10)
    If Hose77_Count > 0 Then totalStr = totalStr & "Напорные рукава 77 мм - " & Hose77_Count & Chr(10)
    If Hose89_Count > 0 Then totalStr = totalStr & "Напорные рукава 89 мм - " & Hose89_Count & Chr(10)
    If Hose110_Count > 0 Then totalStr = totalStr & "Напорные рукава 110 мм - " & Hose110_Count & Chr(10)
    If Hose150_Count > 0 Then totalStr = totalStr & "Напорные рукава 150 мм - " & Hose150_Count & Chr(10)
    If Hose200_Count > 0 Then totalStr = totalStr & "Напорные рукава 200 мм - " & Hose200_Count & Chr(10)
    If Hose250_Count > 0 Then totalStr = totalStr & "Напорные рукава 250 мм - " & Hose250_Count & Chr(10)
    If Hose300_Count > 0 Then totalStr = totalStr & "Напорные рукава 300 мм - " & Hose300_Count & Chr(10)
    If OtherHoses_Count > 0 Then totalStr = totalStr & "Прочие напорные рукава - " & OtherHoses_Count & Chr(10)
    If NapHoses_Lenight > 0 Then totalStr = totalStr & "Длина напорных рукавных линий - " & NapHoses_Lenight & " м " & Chr(10)
    If Hose77NV_Count > 0 Then totalStr = totalStr & "Напорно-всасывающие рукава 77 мм - " & Hose77NV_Count & Chr(10)
    If Hose125NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 125 мм - " & Hose125NV_Count & Chr(10)
    If Hose150NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 150 мм - " & Hose150NV_Count & Chr(10)
    If Hose200NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 200 мм - " & Hose200NV_Count & Chr(10)
    If VsasHoses_Lenight > 0 Then totalStr = totalStr & _
            "Длина всасывающих (напорно-всасывающих) рукавных линий - " & VsasHoses_Lenight & " м " & Chr(10)
    
    If StvB_Count > 0 Then totalStr = totalStr & "Стволов Б - " & StvB_Count & Chr(10)
    If StvA_Count > 0 Then totalStr = totalStr & "Стволов А - " & StvA_Count & Chr(10)
    If StvLaf_Count > 0 Then totalStr = totalStr & "Лафетных стволов - " & StvLaf_Count & Chr(10)
    If StvPen_Count > 0 Then totalStr = totalStr & "Пенных стволов - " & StvPen_Count & Chr(10)
'    If StvGPS_Count > 0 Then totalStr = totalStr & "Пенных стволов стволов - " & StvGPS_Count & Chr(10)
    
    If Razv_Count > 0 Then totalStr = totalStr & "Разветвлений - " & Razv_Count & Chr(10)
    If VS_Count > 0 Then totalStr = totalStr & "Водосборников - " & VS_Count & Chr(10)
    If Collector > 0 Then totalStr = totalStr & "Коллекторов - " & VS_Count & Chr(10)
    If GE_Count > 0 Then totalStr = totalStr & "Гидроэлеваторов - " & GE_Count & Chr(10)
    If PS_Count > 0 Then totalStr = totalStr & "Пеносмесителей - " & PS_Count & Chr(10)
    If PV_Count > 0 Then totalStr = totalStr & "Пенных вставок - " & PV_Count & Chr(10)
    If VsasSetc_Count > 0 Then totalStr = totalStr & "Всасывающих сеток - " & VsasSetc_Count & Chr(10)
    If Kol_Count > 0 Then totalStr = totalStr & "Колонок - " & Kol_Count & Chr(10)
    
    If PG_Count > 0 Then totalStr = totalStr & "Использовано пожарных гидрантов - " & PG_Count & Chr(10)
    If PW_Count > 0 Then totalStr = totalStr & "Использовано пожарных водоемов - " & PW_Count & Chr(10)
    If PK_Count > 0 Then totalStr = totalStr & "Использовано пожарных кранов - " & PK_Count & Chr(10)
    If WaterContainers_Count > 0 Then totalStr = totalStr & "Промежуточных емкостей для воды - " & WaterContainers_Count & Chr(10)
    
    If PW_Count > 0 Then totalStr = totalStr & "Пожарных водоемов - " & PW_Count & Chr(10)

    
    MsgBox totalStr, vbOKOnly, "Анализ насосно-рукавной системы"
End Sub



Public Sub Test()
    GESystemTest Application.ActiveWindow.Selection(1)
End Sub

Public Sub ClearVaraibles()
'Очищаем все переменные
     PA_Count = 0
     MP_Count = 0
     Train_Count = 0
     Ship_Count = 0
    
     Hose38_Count = 0
     Hose51_Count = 0
     Hose66_Count = 0
     Hose77_Count = 0
     Hose89_Count = 0
     Hose110_Count = 0
     Hose150_Count = 0
     Hose200_Count = 0
     Hose250_Count = 0
     Hose300_Count = 0
     OtherHoses_Count = 0
    
     NapHoses_Lenight = 0
     VsasHoses_Lenight = 0
    
     Hose77NV_Count = 0
     Hose125NV_Count = 0
     Hose150NV_Count = 0
     Hose200NV_Count = 0
    
     Razv_Count = 0
     VS_Count = 0
     Collector = 0
     GE_Count = 0
     PS_Count = 0
     VsasSetc_Count = 0
     Kol_Count = 0
     PV_Count = 0
    
     StvA_Count = 0
     StvB_Count = 0
     StvLaf_Count = 0
     StvPen_Count = 0
     StvGPS_Count = 0
     
     PodOut = 0
     PodIn = 0
     HosesValue = 0
     WaterValue = 0
'     WaterContainers_Value = 0
    
     PG_Count = 0
     PW_Count = 0
     PK_Count = 0
     WaterContainers_Count = 0
     
     PW_Count = 0
     PW_Value = 0
     WaterOpen_Value = 0
End Sub

'----------------------------------------Служебные Функции-----------------------------------------------
Public Sub NRS_Analize()
'Процедура анализа насосно-рукавной системы
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer

    On Error GoTo EX

'---Очищаем все переменные
    ClearVaraibles

'---Перебираем все фигуры и в случае если фигура является фигурой ГраФиС анализируем её
    For Each vsO_Shape In shapesInNRS
        If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then 'Является ли фигура фигурой ГраФиС
            If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then  'Проверяем версию фигуры
                vsi_ShapeIndex = vsO_Shape.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
                
                '---Проверяем маневерннеость (если фигура маневренная  выходим)
                If IsNotManeuwer(vsO_Shape) Then
                    
                
                    '---Общие свойства (характерные для нескольких видов фигур, но именно ГраФиС)
                    '---Расход из лафетного ствола
                    If vsO_Shape.CellExists("User.GFS_OutLafet", 0) = True Then
                        PodOut = PodOut + vsO_Shape.Cells("User.GFS_OutLafet").Result(visNumber)
                    End If

                
                    Select Case vsi_ShapeIndex
                    '---Пожарные автомобили-----------
                        Case Is = indexPers.ipAC 'Автоцистерны
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipANR 'АНР
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipPNS 'ПНС
                            PA_Count = PA_Count + 1
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAA 'AA
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAV 'АВ
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAKT 'АКТ
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAGVT 'АГВТ
                            PA_Count = PA_Count + 1
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 2 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAR 'АР
                            PA_Count = PA_Count + 1
                        Case Is = indexPers.ipACL 'АЦЛ
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipACKP 'АЦКП
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                        Case Is = indexPers.ipAPP 'АПП
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 'Засчитываем водосборник на а/м
                    
                    '---Прочая пожарная техника----------
                        Case Is = indexPers.ipPoezd 'Поезда
                            Train_Count = Train_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = indexPers.ipMotoPump 'Мотопомпы
                            MP_Count = MP_Count + 1
                        Case Is = indexPers.ipKorabl 'Корабль
                            Ship_Count = Ship_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                        Case Is = indexPers.ipKater 'Катер
                            Ship_Count = Ship_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                        Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
                        Case Is = indexPers.ipTank 'Танки
            
                    '---ПТВ-----------------------------------
                        Case Is = indexPers.ipStvolRuch 'Ручной водяной
                            If vsO_Shape.Cells("User.DiameterIn").Result(visNumber) <= 50 Then
                                StvB_Count = StvB_Count + 1
                            Else
                                StvA_Count = StvA_Count + 1
                            End If
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolRuchPena 'Ручной пенный
                            StvPen_Count = StvPen_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafVoda 'Лафетный водяной
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafPena 'Лафетный пенный
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafVoda 'Лафетный водяной возимый
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipRazvetvlenie 'Разветвление
                            Razv_Count = Razv_Count + 1
                        Case Is = indexPers.ipVodosbornik 'Vodosbornik
                            VS_Count = VS_Count + 1
                        Case Is = indexPers.ipKollector 'Collector
                            Collector = Collector + 1
                        Case Is = indexPers.ipPenopodemnik 'Пеноподъемник
                        Case Is = indexPers.ipGrebenkaPGenerat 'Гребенка
                        Case Is = indexPers.ipPennayaVstavka 'Пенная вставка
                            PV_Count = PV_Count + 1
    
            
                    '---Водоснабжение
                        Case Is = indexPers.ipPG 'ПГ
                            PG_Count = PG_Count + 1
                        Case Is = indexPers.ipPW 'ПВ
                            PW_Count = PW_Count + 1
                        Case Is = indexPers.ipPK 'ПК
                            PK_Count = PK_Count + 1
                        Case Is = indexPers.ipOtkritiyVodoistochnik 'Водоем
                        Case Is = indexPers.ipPirs 'Пирс
                        Case Is = indexPers.ipBashna 'Башня
                            
                    '---Забор воды
                        Case Is = indexPers.ipHydroelevator 'Гидроэлеватор
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "Да" Then
                                GE_Count = GE_Count + 1
                                GetWaterSourceData vsO_Shape
                            End If
                        Case Is = indexPers.ipPenosmesitelPerenosn 'Пеносмеситель
                            PS_Count = PS_Count + 1
                        Case Is = indexPers.ipVsasLineWithSetk  'Всасывающая линия с сеткой
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "Да" Then
                                VsasSetc_Count = VsasSetc_Count + 1
    '                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                                GetWaterSourceData vsO_Shape
                            End If
                        Case Is = indexPers.ipKolonka 'Колонка
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "Да" Then
                                Kol_Count = Kol_Count + 1
                                PodIn = PodIn + vsO_Shape.Cells("Prop.FlowCurrent").Result(visNumber)
                            End If
                        Case Is = indexPers.ipEmkost 'Емкости для воды
                            WaterContainers_Count = WaterContainers_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.WaterContainerValue").Result(visNumber) * _
                                    vsO_Shape.Cells("Prop.OstKoeff").Result(visNumber)
                       
                    '---Линии
                        Case Is = indexPers.ipRukavLineNapor 'Напорная линия
                            If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "Нет" _
                                And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                                Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
                                    Case Is = 38
                                        Hose38_Count = Hose38_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 51
                                        Hose51_Count = Hose51_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 66
                                        Hose66_Count = Hose66_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 77
                                        Hose77_Count = Hose77_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 89
                                        Hose89_Count = Hose89_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 110
                                        Hose110_Count = Hose110_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 150
                                        Hose150_Count = Hose150_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 200
                                        Hose200_Count = Hose200_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 250
                                        Hose250_Count = Hose250_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 300
                                        Hose300_Count = Hose300_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                        
                                End Select
                                '---Длина линии
                                If vsO_Shape.CellExists("User.TotalLenight", 0) = True Then 'Имеет ли фигура ячейку "User.TotalLenight"
                                    NapHoses_Lenight = NapHoses_Lenight + vsO_Shape.Cells("User.TotalLenight")
                                Else
                                    NapHoses_Lenight = NapHoses_Lenight + vsO_Shape.Cells("Prop.LineLenightHose")
                                End If
                                '---Объем воды в рукаве
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
                            If vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                                OtherHoses_Count = OtherHoses_Count + 1
                                '---Объем воды в рукаве
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
                        Case Is = indexPers.ipRukavLineVsas 'Всасывающая линия или напорно-всасывающая
                            If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "Нет" _
                                And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                                Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
                                    Case Is = 77
                                        Hose77NV_Count = Hose77NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 125
                                        Hose125NV_Count = Hose125NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 150
                                        Hose150NV_Count = Hose150NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 200
                                        Hose200NV_Count = Hose200NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                End Select
                                '---Длина линии
                                If vsO_Shape.CellExists("User.TotalLenight", 0) = True Then 'Имеет ли фигура ячейку "User.TotalLenight"
                                    VsasHoses_Lenight = VsasHoses_Lenight + vsO_Shape.Cells("User.TotalLenight")
                                Else
                                    VsasHoses_Lenight = VsasHoses_Lenight + vsO_Shape.Cells("Prop.LineLenightHose")
                                End If
                                '---Объем воды в рукаве
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
    
                    End Select
                End If
            End If
        End If
    Next vsO_Shape



Exit Sub
EX:
    SaveLog Err, "NRS_Analize"
    
End Sub

Private Function IsNotManeuwer(ByRef shp As Visio.Shape) As Boolean
'Функция возвращает Истину, если фигура не маневренная, или такое свойство у нее вообще отсутствует
    If shp.CellExists("Actions.MainManeure.Checked", 0) = True Then
        If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
            IsNotManeuwer = True
        Else
            IsNotManeuwer = False
        End If
        Exit Function
    End If
IsNotManeuwer = True
End Function

Private Sub GetWaterSourceData(ByRef shp As Visio.Shape)
'Получаем сведения о задействованном водоеме от фигур всасывающей линии с сеткой или гидроэлеватора
Dim shpWS As Visio.Shape
Dim shpWSIndex As Long
Dim tmpVal As Long
    
    shpWSIndex = CellVal(shp, "User.WSShapeID")
    If shpWSIndex > 0 Then
        Set shpWS = Application.ActivePage.Shapes.ItemFromID(shpWSIndex)
        If IsGFSShapeWithIP(shpWS, indexPers.ipPW) Then
            PW_Count = PW_Count + 1
            
            tmpVal = CellVal(shpWS, "Prop.WaterValue") * 1000
            PW_Value = PW_Value + tmpVal
        End If
        If IsGFSShapeWithIP(shpWS, indexPers.ipOtkritiyVodoistochnik) Then
            If CellVal(shpWS, "Prop.Type", visUnitsString) = "Неограниченный запас" Then
                PodIn = PodIn + CellVal(shp, "Prop.PodIn")
                PodIn = PodIn + CellVal(shp, "Prop.PodFromOuter")
            Else
                tmpVal = CellVal(shpWS, "Prop.Value") * 1000 * 0.9  'С учетом коэффициента использования
                WaterOpen_Value = WaterOpen_Value + tmpVal
            End If
        End If
    End If
    
End Sub












