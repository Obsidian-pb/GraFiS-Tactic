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
Public GE_Count As Integer
Public PS_Count As Integer
Public VsasSetc_Count As Integer
Public Kol_Count As Integer

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





Public Sub GESystemTest(ShpObj As Visio.Shape)
'Основная процедура получения сведений о насосно-рукавной системе
    
    On Error GoTo ex
    
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
ex:
    Set shapesInNRS = Nothing
End Sub

Private Sub GetTechShapeForGESystem(ByRef shp As Visio.Shape)
'Заполняем коллекцию фигур соединенных в НРС
Dim con As Connect
Dim sideShp As Visio.Shape

    For Each con In shp.Connects
        If Not IsShapeAllreadyChecked(con.ToSheet) Then
            shapesInNRS.Add con.ToSheet
            GetTechShapeForGESystem con.ToSheet
        End If
    Next con
    For Each con In shp.FromConnects
        If Not IsShapeAllreadyChecked(con.FromSheet) Then
            shapesInNRS.Add con.FromSheet
            GetTechShapeForGESystem con.FromSheet
        End If
    Next con

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
    
    If PodOut > 0 Then totalStr = totalStr & "Общий расход системы - " & PodOut & "л/с" & Chr(10)
    If PodIn > 0 Then totalStr = totalStr & "Общий забор воды - " & PodIn & "л/с" & Chr(10)
    If HosesValue > 0 Then totalStr = totalStr & "Объем воды в рукавах - " & HosesValue & "л" & Chr(10)
    If WaterValue > 0 Then totalStr = totalStr & "Запас воды в емкостях - " & WaterValue & "л" & Chr(10)
    
    If PodOut > PodIn Then
        Dim FlowOut As Double 'Скорость убывания жидкости
        Dim DischargeTime As Double
        FlowOut = PodOut - PodIn
        DischargeTime = ((WaterValue - HosesValue) / FlowOut) / 60
        totalStr = totalStr & "Возможное время работы системы - " & _
                 Int(DischargeTime) & ":" & Int((DischargeTime - Int(DischargeTime)) * 60) _
                 & Chr(10)
    Else
        totalStr = totalStr & "Возможное время работы системы - бесконечно" & Chr(10)
    End If
    
    If PA_Count > 0 Then totalStr = totalStr & "Пожарных автомобилей - " & PA_Count & Chr(10)
    If MP_Count > 0 Then totalStr = totalStr & "Пожарных мотопомп - " & MP_Count & Chr(10)
    
    If Hose51_Count > 0 Then totalStr = totalStr & "Напорные рукава 51мм - " & Hose51_Count & Chr(10)
    If Hose66_Count > 0 Then totalStr = totalStr & "Напорные рукава 66мм - " & Hose66_Count & Chr(10)
    If Hose77_Count > 0 Then totalStr = totalStr & "Напорные рукава 77мм - " & Hose77_Count & Chr(10)
    If Hose89_Count > 0 Then totalStr = totalStr & "Напорные рукава 89мм - " & Hose89_Count & Chr(10)
    If Hose110_Count > 0 Then totalStr = totalStr & "Напорные рукава 110мм - " & Hose110_Count & Chr(10)
    If Hose150_Count > 0 Then totalStr = totalStr & "Напорные рукава 150мм - " & Hose150_Count & Chr(10)
    If Hose200_Count > 0 Then totalStr = totalStr & "Напорные рукава 200мм - " & Hose200_Count & Chr(10)
    If Hose250_Count > 0 Then totalStr = totalStr & "Напорные рукава 250мм - " & Hose250_Count & Chr(10)
    If Hose300_Count > 0 Then totalStr = totalStr & "Напорные рукава 300мм - " & Hose300_Count & Chr(10)
    If OtherHoses_Count > 0 Then totalStr = totalStr & "Прочие напорные рукава - " & OtherHoses_Count & Chr(10)
    If NapHoses_Lenight > 0 Then totalStr = totalStr & "Длина напорных рукавных линий - " & NapHoses_Lenight & "м " & Chr(10)
    If Hose77NV_Count > 0 Then totalStr = totalStr & "Напорно-всасывающие рукава 77мм - " & Hose77NV_Count & Chr(10)
    If Hose125NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 125мм - " & Hose125NV_Count & Chr(10)
    If Hose150NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 150мм - " & Hose150NV_Count & Chr(10)
    If Hose200NV_Count > 0 Then totalStr = totalStr & "Всасывающие рукава 200мм - " & Hose200NV_Count & Chr(10)
    If VsasHoses_Lenight > 0 Then totalStr = totalStr & _
            "Длина всасывающих (напорно-всасывающих) рукавных линий - " & VsasHoses_Lenight & "м " & Chr(10)
    
    If StvB_Count > 0 Then totalStr = totalStr & "Стволов Б - " & StvB_Count & Chr(10)
    If StvA_Count > 0 Then totalStr = totalStr & "Стволов А - " & StvA_Count & Chr(10)
    If StvLaf_Count > 0 Then totalStr = totalStr & "Лафетных стволов - " & StvLaf_Count & Chr(10)
    If StvPen_Count > 0 Then totalStr = totalStr & "Пенных стволов - " & StvPen_Count & Chr(10)
'    If StvGPS_Count > 0 Then totalStr = totalStr & "Пенных стволов стволов - " & StvGPS_Count & Chr(10)
    
    If Razv_Count > 0 Then totalStr = totalStr & "Разветвлений - " & Razv_Count & Chr(10)
    If GE_Count > 0 Then totalStr = totalStr & "Гидроэлеваторов - " & GE_Count & Chr(10)
    If PS_Count > 0 Then totalStr = totalStr & "Пеносмесителей - " & PS_Count & Chr(10)
    If VsasSetc_Count > 0 Then totalStr = totalStr & "Всасывающих сеток - " & VsasSetc_Count & Chr(10)
    If Kol_Count > 0 Then totalStr = totalStr & "Колонок - " & Kol_Count & Chr(10)
    
    If PG_Count > 0 Then totalStr = totalStr & "Использовано пожарных гидрантов - " & PG_Count & Chr(10)
    If PW_Count > 0 Then totalStr = totalStr & "Использовано пожарных водоемов - " & PW_Count & Chr(10)
    If PK_Count > 0 Then totalStr = totalStr & "Использовано пожарных кранов - " & PK_Count & Chr(10)
    If WaterContainers_Count > 0 Then totalStr = totalStr & "Промежуточных емкостей для воды - " & WaterContainers_Count & Chr(10)
    
    MsgBox totalStr, vbOKOnly, "Анализ насосно-рукавной системы"
End Sub



Public Sub Test()
    GESystemTest Application.ActiveWindow.Selection(1)
End Sub

Public Sub ClearVaraibles()
'Очищаем все переменные
     PA_Count = 0
     MP_Count = 0
    
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
     GE_Count = 0
     PS_Count = 0
     VsasSetc_Count = 0
     Kol_Count = 0
    
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
     
End Sub

'----------------------------------------Служебные Функции-----------------------------------------------
Public Sub NRS_Analize()
'Процедура анализа насосно-рукавной системы
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer

    On Error GoTo ex

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
                    '---РАсход из лафетного ствола
                    If vsO_Shape.CellExists("User.GFS_OutLafet", 0) = True Then
                        PodOut = PodOut + vsO_Shape.Cells("User.GFS_OutLafet").Result(visNumber)
                    End If

                
                    Select Case vsi_ShapeIndex
                    '---Пожарные автомобили-----------
                        Case Is = 1 'Автоцистерны
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = 2 'АНР
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = 8 'ПНС
                            PA_Count = PA_Count + 1
                        Case Is = 9 'AA
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = 10 'АВ
                            PA_Count = PA_Count + 1
                        Case Is = 11 'АКТ
                            PA_Count = PA_Count + 1
                        Case Is = 13 'АГВТ
                            PA_Count = PA_Count + 1
                        Case Is = 20 'АР
                            PA_Count = PA_Count + 1
                        Case Is = 161 'АЦЛ
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = 162 'АЦКП
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = 163 'АПП
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                    
                    '---Прочая пожарная техника----------
                        Case Is = 24 'Поезда
                        Case Is = 28 'Мотопомпы
                            MP_Count = MP_Count + 1
                        Case Is = 30 'Корабль
                        Case Is = 31 'Катер
                        Case Is = 73 'Машины на гусеничном ходу
                        Case Is = 74 'Танки
            
                    '---ПТВ-----------------------------------
                        Case Is = 34 'Ручной водяной
                            If vsO_Shape.Cells("User.DiameterIn").Result(visNumber) = 50 Then
                                StvB_Count = StvB_Count + 1
                            Else
                                StvA_Count = StvA_Count + 1
                            End If
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = 35 'Ручной пенный
                            StvPen_Count = StvPen_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = 36 'Лафетный водяной
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = 37 'Лафетный пенный
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = 39 'Лафетный водяной возимый
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = 42 'Разветвление
                            Razv_Count = Razv_Count + 1
                        Case Is = 45 'Пеноподъемник
                        Case Is = 22 'Гребенка
    
    
            
                    '---Водоснабжение
                        Case Is = 50 'ПГ
                            PG_Count = PG_Count + 1
                        Case Is = 51 'ПВ
                            PW_Count = PW_Count + 1
                        Case Is = 52 'ПК
                            PK_Count = PK_Count + 1
                        Case Is = 53 'Водоем
                        Case Is = 54 'Пирс
                        Case Is = 56 'Башня
                            
                    '---Забор воды
                        Case Is = 40 'Гидроэлеватор
                            GE_Count = GE_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodFromOuter").Result(visNumber)
                        Case Is = 41 'Пеносмеситель
                            PS_Count = PS_Count + 1
                        Case Is = 88  'Всасывающая линия с сеткой
                            VsasSetc_Count = VsasSetc_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                        Case Is = 72 'Колонка
                            Kol_Count = Kol_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.FlowCurrent").Result(visNumber)
                        Case Is = 190 'Емкости для воды
                            WaterContainers_Count = WaterContainers_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.WaterContainerValue").Result(visNumber) * _
                                    (1 - vsO_Shape.Cells("Prop.OstKoeff").Result(visNumber))
                            
                    '---Линии
                        Case Is = 100 'Напорная линия
                            If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "Нет" _
                                And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                                Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
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
                        Case Is = 101 'Всасывающая линия или напорно-всасывающая
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
ex:
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
