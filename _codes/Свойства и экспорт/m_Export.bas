Attribute VB_Name = "m_Export"
Option Explicit

'Разделеитель строк в командах и информации
Const delimiter = " | "
'Знак границы маркера
Const mChar = "$"
'Путь и названия шаблонов документов
Const pathNameDonesenie = "Templates\Donesenie.dot"      'Донесение о пожаре
Const pathNameKBD = "Templates\BD_Card.dot"      'Карточка боевых действий




'------------Процедуры экспорта данных формы--------------------------
Public Sub ExportToWord_Donesenie()
'Экспортируем данные формы в документ Word - в Донесение
Dim wrd As Object
Dim wrdDoc As Object
Dim path As String

Dim gfsShapes As Collection
Dim gettedDate As Variant
Dim gettedCol As Collection
Dim gettedTxt As String


    'Предварительно запускаем процедуру исправления английских C на русские С:
'    fixAllGFSShapesC

    'Определяем путь к шаблону документа
    path = ThisDocument.path & pathNameDonesenie
    'Создаем новый документ Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add(path)
    wrdDoc.Activate
    
    
    
    
    
    'Блок кода отвечающего за экспорт данных
    'В случае ошибки пропускаем строку кода:
'    On Error Resume Next
    
    'Формируем основную коллекцию фигур ГраФиС:
    Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
    'Процедуры экспорта:
    '---Основное
'        gettedTxt = cellVal(gfsShapes, "Prop.NP_Name", visUnitsString, "")
        SetData wrd, "НП", cellVal(gfsShapes, "Prop.NP_Name", visUnitsString, "")                       'Населенный пункт
        SetData wrd, "ДЗФИО", cellVal(gfsShapes, "Prop.PersonCreate", visUnitsString, "")                      'должность, звание, фамилия, имя, отчество (при наличии)
        SetData wrd, "Наименование", cellVal(gfsShapes, "Prop.ObjectName", visUnitsString, "")          'Наименование
        SetData wrd, "Принадлежность", cellVal(gfsShapes, "Prop.Affiliation", visUnitsString, "")       'Принадлежность объекта
        SetData wrd, "Адрес", cellVal(gfsShapes, "Prop.Address", visUnitsString, "")                    'Адрес организации
        SetData wrd, "МестоПожара", cellVal(gfsShapes, "Prop.FireStartPlace", visUnitsString, "")              'Место возникновения пожара
        SetData wrd, "Заявитель", cellVal(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString, "")               'Фамилия, имя, отчество (при наличии) лица, обнаружившего пожар и способ сообщения о нем в пожарную охрану
        SetData wrd, "Заявитель_т", cellVal(gfsShapes, "Prop.CallPhone", visUnitsString, "")            'Номер телефона заявителя
        
        
        
        
'Сделать это:
    '---Даты и время
        'Дата пожара:
        gettedDate = CDate(A.Result("FireTime"))
        If gettedDate > 0 Then
            SetData wrd, "П_День", Format(gettedDate, "DD")                             'День возникновения пожара
            SetData wrd, "П_Месяц", Split(Format(gettedDate, "DD MMMM"), " ")(1)        'Месяц возникновения пожара
            SetData wrd, "П_Год", Format(gettedDate, "YY")                              'Месяц возникновения пожара
        End If
        'Дата Обнаружения:
        gettedDate = CDate(A.Result("FindTime"))
        If gettedDate > 0 Then
            SetData wrd, "Обн_Час", Format(gettedDate, "HH")                           'Час обнаружения пожара
            SetData wrd, "Обн_Мин", Format(gettedDate, "NN")                           'Минута обнаружения пожара
        End If
        'Дата сообщения:
        gettedDate = CDate(A.Result("InfoTime"))
        If gettedDate > 0 Then
            SetData wrd, "Сооб_Дата", Format(gettedDate, "DD.MM.YYYY")                'Дата сообщения о пожаре
            SetData wrd, "Сооб_Час", Format(gettedDate, "HH")                           'Час сообщения о пожаре
            SetData wrd, "Сооб_Мин", Format(gettedDate, "NN")                           'Минута сообщения о пожаре
        End If
        'Время прибытия первого подразделения:
        gettedDate = CDate(A.Result("FirstArrivalTime"))
        If gettedDate > 0 Then
            SetData wrd, "1Ств_Дата", Format(gettedDate, "DD.MM.YYYY")
            SetData wrd, "1Подр_Час", Format(gettedDate, "HH")                           'Час сообщения о пожаре
            SetData wrd, "1Подр_Мин", Format(gettedDate, "NN")                           'Минута сообщения о пожаре
        End If
        
        'Дата и время подачи первого ствола
        gettedDate = CDate(A.Result("FirstStvolTime"))
        If gettedDate > 0 Then
        SetData wrd, "1Ств_Час", Format(gettedDate, "HH")
        SetData wrd, "1Ств_Мин", Format(gettedDate, "NN")
        End If
        
        'Дата и время локализации пожара
        gettedDate = CDate(A.Result("LocalizationTime"))
        If gettedDate > 0 Then
        SetData wrd, "Лок_Дата", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "Лок_Час", Format(gettedDate, "HH")
        SetData wrd, "Лок_Мин", Format(gettedDate, "NN")
        End If
        
        'Дата и время ликвидации открытого горения
        gettedDate = CDate(A.Result("LOGTime"))
        If gettedDate > 0 Then
        SetData wrd, "ЛОГ_Дата", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "ЛОГ_Час", Format(gettedDate, "HH")
        SetData wrd, "ЛОГ_Мин", Format(gettedDate, "NN")
        End If
        
        'Дата и время ликвидации последствий пожара
        gettedDate = CDate(A.Result("LPPTime"))
        If gettedDate > 0 Then
        SetData wrd, "ЛПП_Дата", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "ЛПП_Час", Format(gettedDate, "HH")
        SetData wrd, "ЛПП_Мин", Format(gettedDate, "NN")
        End If
        
        'Обстановка на момент прибытия
        Set gettedCol = A.GetGFSShapesAnd("User.IndexPers:604;Prop.SituationKind:на момент прибытия")
        If gettedCol.Count > 0 Then
            SetData wrd, "Обст_Приб", cellVal(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
        End If
        
        'Количество звеньев ГДЗС
        gettedDate = CDate(A.Result("GDZSChainsCountWork"))
        If gettedDate > 0 Then
        SetData wrd, "ГДЗС_Кол", Format(A.Result("GDZSChainsCountWork"))
        End If
        
        'Число участников тушения
        gettedDate = CDate(A.Result("PersonnelHave"))
        If gettedDate > 0 Then
        SetData wrd, "Уч_Туш", Format(A.Result("PersonnelHave"))
        End If
        
        
        'Количество основных и специальных отделений
        SetData wrd, "Осн_Спец", "Основных ПА: " & A.Result("MainOverallHave") & ", Специальных ПА:" & A.Result("SpecialPAHave")
        
        
        gettedTxt = ""
        gettedDate = A.Result("StvolWHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "вода, "
        End If
        
        gettedDate = A.Result("StvolFoamHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "Пена, "
         End If
            
            gettedDate = A.Result("StvolGasHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "Газ. "
        End If
        
        SetData wrd, "ОВ?", gettedTxt
        
        
        'Погибло людей
        gettedTxt = cellVal(gfsShapes, "Prop.HumansDie", visUnitsString, "")
'        SetData wrd, "200", "Погибло людей: " & Split(gettedTxt, "/")(0) & "в том числе детей: " & Split(gettedTxt, "/")(1) & "работников ПО: " & Split(gettedTxt, "/")(2)
        SetData wrd, "200", Split(gettedTxt, "/")(0)
        SetData wrd, "200Д", Split(gettedTxt, "/")(1)
        SetData wrd, "200ПО", Split(gettedTxt, "/")(2)
        
        
        
         'Травмировано людей
        gettedTxt = cellVal(gfsShapes, "Prop.HumansInjured", visUnitsString, "")
        SetData wrd, "300", Split(gettedTxt, "/")(0)
        SetData wrd, "300Д", Split(gettedTxt, "/")(1)
        SetData wrd, "300ПО", Split(gettedTxt, "/")(2)
        
        'Информация о погибших и травмированных
        Set gettedCol = GetVictims
        SetData wrd, "200Св", gettedCol(1)
        SetData wrd, "300Св", gettedCol(2)
        
        'Уничтожено/повреждено
        '---строений
        gettedTxt = cellVal(gfsShapes, "Prop.ConstructionsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Стр", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Стр", Split(gettedTxt, "/")(1)
        '---Квартир
        gettedTxt = cellVal(gfsShapes, "Prop.FlatsAffected", visUnitsString, "")
        SetData wrd, "Уничт_ЖК", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_ЖК", Split(gettedTxt, "/")(1)
        '---Комнат
        gettedTxt = cellVal(gfsShapes, "Prop.RoomsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Комн", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Комн", Split(gettedTxt, "/")(1)
        '---Площади
        gettedTxt = cellVal(gfsShapes, "Prop.SquareAffected", visUnitsString, "")
        SetData wrd, "Уничт_Пл", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Пл", Split(gettedTxt, "/")(1)
        '---Техники
        gettedTxt = cellVal(gfsShapes, "Prop.TechnicsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Тех", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Тех", Split(gettedTxt, "/")(1)
        '---сельхоз культур
        gettedTxt = cellVal(gfsShapes, "Prop.AgricultureAffected", visUnitsString, "")
        SetData wrd, "Уничт_СХ", ClearString(gettedTxt)
        '---сельхоз животных
        gettedTxt = cellVal(gfsShapes, "Prop.CattleAffected", visUnitsString, "")
        SetData wrd, "200СХ", ClearString(gettedTxt)
        'Спасено
        '---людей
        gettedTxt = cellVal(gfsShapes, "Prop.Saved", visUnitsString, "")
        SetData wrd, "Спас_Л", Split(gettedTxt, "/")(0)
        '---техники
        SetData wrd, "Спас_Т", Split(gettedTxt, "/")(1)
        '---голов скота
        SetData wrd, "Спас_Ск", Split(gettedTxt, "/")(2)
        
        
        'Перечень подразделений
        Set gettedCol = GetUniqueVals(gfsShapes, "Prop.Unit", , "-", "-")
        SetData wrd, "Подразделения", StrColToStr(gettedCol, ", ")
        'Тип, количество и принадлежность пожарной техники
        SetData wrd, "Техника", GetTechniks(gettedCol)
        'Количество и вид поданных стволов
        SetData wrd, "КолВид_Ст", GetReadyStringA("StvolWBHave", "Стволов Б -", ", ") & _
                     GetReadyStringA("StvolWAHave", "Стволов А -", ", ") & _
                     GetReadyStringA("StvolWLHave", "Лафетных стволов -", ", ") & _
                     GetReadyStringA("StvolFoamHave", "Пенных стволов -", ", ")
                     
        
        
        
        SetData wrd, "ПожАвт", cellVal(gfsShapes, "Prop.FireAutomatics", visUnitsString, "")
        SetData wrd, "Осн_Спец", "Основных ПА: " & A.Result("MainOverallHave") & ", Специальных ПА:" & A.Result("SpecialPAHave")
        
        
        'Обстоятельства усложняющие обстановку
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesRize", visUnitsString, "")
        SetData wrd, "Обст", ClearString(gettedTxt)
        
        'Силы и средства применявшиеся при тушении
        gettedTxt = cellVal(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "СиС", ClearString(gettedTxt)
        
        'Водоисточники
        gettedTxt = cellVal(gfsShapes, "Prop.WaterSources", visUnitsString, "")
        SetData wrd, "ВидВодоист", ClearString(gettedTxt)
        
        
        
'Сделать это:
    '---Очищаем все незаполненные маркеры
'        clearLostMarkers wrd

        
        
End Sub


Public Sub ExportToWord_KBD()
'Экспортируем данные формы в документ Word - в карточку боевых действий
Dim wrd As Object
Dim wrdDoc As Object
Dim path As String

Dim gfsShapes As Collection
Dim gettedDate As Variant
Dim gettedCol As Collection
Dim gettedTxt As String



    'Определяем путь к шаблону документа
    path = ThisDocument.path & pathNameKBD
    'Создаем новый документ Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add(path)
    wrdDoc.Activate
    
    
    
    
    
    'Блок кода отвечающего за экспорт данных
    'В случае ошибки пропускаем строку кода:
'    On Error Resume Next
    
    'Формируем основную коллекцию фигур ГраФиС:
    Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
    'Процедуры экспорта:
        'Вызов №
        gettedTxt = cellVal(gfsShapes, "Prop.FireRank", visUnitsString)
        SetData wrd, "Вызов", gettedTxt
        'Подразделение
        gettedTxt = cellVal(gfsShapes, "Prop.ThisDocUnit", visUnitsString)
        SetData wrd, "Подр", gettedTxt
        'Дата пожара
        gettedDate = CDate(A.Result("FireTime"))
        SetData wrd, "П_Дата", Format(gettedDate, "DD.MM.YYYY")
        'Наименование организации (объекта), его ведомственная принадлежность (форма собственности), адрес
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectName", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.OrgPrinadl", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.OrgPropertyType", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.Address", visUnitsString)
        SetData wrd, "Инф_Орг", gettedTxt
        'размеры в плане, этажность
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectWidth", visUnitsString)
        gettedTxt = gettedTxt & "х" & cellVal(gfsShapes, "Prop.ObjectLenight", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectFloorCount", visUnitsString) & " этажей"
        SetData wrd, "Хар_Орг1", gettedTxt
        'конструктивные особенности, степень огнестойкости категория производства
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectConstructions", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectSO", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectCP", visUnitsString)
        SetData wrd, "Хар_Орг2", gettedTxt
        'Кем охраняется организация (объект), кто обнаружил пожар
        gettedTxt = cellVal(gfsShapes, "Prop.Guard", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString)
        SetData wrd, "Охр", gettedTxt
        'Время/площадь
        '---возникновения пожара
            gettedDate = CDate(A.Result("FireTime"))
            SetData wrd, "Возн_Вр", Format(gettedDate, "HH:NN")
            SetData wrd, "Возн_Пл", "0"
        '---обнаружения пожара
            gettedDate = CDate(A.Result("FindTime"))
            SetData wrd, "Обн_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Обн_Пл", gettedTxt
        '---сообщения пожара
            gettedDate = CDate(A.Result("InfoTime"))
            SetData wrd, "Сооб_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Сооб_Пл", gettedTxt
        '---выезда караула
            gettedDate = DateAdd("n", 1, CDate(A.Result("InfoTime")))
            SetData wrd, "Выезд_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Выезд_Пл", gettedTxt
        '---прибытия на пожар
            gettedDate = CDate(A.Result("FirstArrivalTime"))
            SetData wrd, "1Подр_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Приб_Пл", gettedTxt
        '---подачи первого ствола
            gettedDate = CDate(A.Result("FirstStvolTime"))
            SetData wrd, "1Ств_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "1Ств_Пл", gettedTxt
        '---вызова дополнительных сил
'            gettedDate = CDate(A.Result("FireTime"))
            SetData wrd, "ДопСил_Вр", "---"
            SetData wrd, "ДопСил_Пл", "---"
        '---локализации
            gettedDate = CDate(A.Result("LocalizationTime"))
            SetData wrd, "Лок_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Лок_Пл", gettedTxt
        '---ликвидации открытого горения
            gettedDate = CDate(A.Result("LOGTime"))
            SetData wrd, "ЛОГ_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "ЛОГ_Пл", gettedTxt
        '---ликвидации
            gettedDate = CDate(A.Result("LPPTime"))
            SetData wrd, "Ликв_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Ликв_Пл", gettedTxt
        '---возвращения в часть
            gettedDate = CDate(A.Result("FireEndTime"))
            SetData wrd, "Возв_Вр", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "Возв_Пл", gettedTxt
        'Возвращаем анализ полного набора данных
        Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
        'Водоснабжение
        gettedTxt = cellVal(gfsShapes, "Prop.WaterSources", visUnitsString, " ")
        SetData wrd, "Водоснабжение", ClearString(gettedTxt)
        'Способы подачи воды:
        '---
            
            
            
            
        
        'Обстановка на момент прибытия
        Set gettedCol = A.GetGFSShapesAnd("User.IndexPers:604;Prop.SituationKind:на момент прибытия")
        If gettedCol.Count > 0 Then
            SetData wrd, "Обст", cellVal(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
        End If
        'Оценка действий
        gettedTxt = GetMarkStr(gfsShapes)
        SetData wrd, "ОценкаДействий", gettedTxt
        
        'Штаб
        gettedTxt = Format(cellVal(gfsShapes, "Prop.StabCreationTime", visDate), "HH:NN") & ". " & GetStabMembers
        SetData wrd, "ШТАБ", gettedTxt
        
        'БУ/СТП - время, задачи участков (секторов) тушения пожара
        gettedTxt = GetBUSTPString
        SetData wrd, "БУ/СТП", gettedTxt
        
        'Обстоятельства, способствующие развитию пожара
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesRize", visUnitsString, " ")
        SetData wrd, "Обст", gettedTxt
        'Обстоятельства, усложняющие обстановку
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesComplicate", visUnitsString, " ")
        SetData wrd, "Слож", gettedTxt
        
        'Силы и средства применявшиеся при тушении
        gettedTxt = cellVal(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "СиС", ClearString(gettedTxt)
        'С использованием техники организаций (объектов)
        '---Не реализовано
        SetData wrd, "ТехОрг", "---"
        '---С использованием сил и средств опорных пунктов тушения крупных пожаров
        SetData wrd, "СиСОП", "---"
        
        'ГДЗС 46,90 (ДАСВ, ДАСК)
        Set gettedCol = A.GetGFSShapes("User.IndexPers:46;User.IndexPers:90")
        If gettedCol.Count > 0 Then
            SetData wrd, "ГДЗС", A.Result("GDZSChainsCountWork") & " звеньев, " & A.Result("GDZSMansCountWork") & " газодымозащитинков"
        End If
        '1 звено
        '---Не реализовано
        SetData wrd, "1ЗВ", "---"
        '2 звена и более
        '---Не реализовано
        SetData wrd, "2ЗВ", "---"
        
        'С какими службами было организовано взаимодействие
        gettedTxt = GetServicesCommunications
        SetData wrd, "Взаим_Сл", gettedTxt

        
        
        
        
        
End Sub
'Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String, _
'                    Optional ByVal ignore As Variant = 0, Optional ByVal ifIgnore As Variant = " ")
Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String)
'Замена маркера в тексте текстом из результата анализа
    
    On Error GoTo ex
    
    'Если пришедшие данные необходимо проигнорировать (например, в случае, если дата равна 0), приравниваем txt = ifIgnore
'    If txt = ignore Then txt = ifIgnore
    
    'Добавляем к имени маркера граничные знаки: "markerName"=>"$markerName$"
    markerName = mChar & markerName & mChar
    
    'Собственно заменяем все одноименные маркеры текстом txt
    With wrd
        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
        With .Selection.Find
            .Text = markerName
            .Replacement.Text = txt
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Selection.Find.Execute Replace:=2
    End With
    
ex:
    
End Sub

Private Sub clearLostMarkers(ByRef wrd As Object, Optional ByVal defaultVal As String = "     ")
Dim markers() As String
Dim marker As String
Dim i As Integer

    markers = Split("Обст_Приб;СиС;ВидВодоист;200Св;300Д;300Св;ЖК;Комнат;Поэт_Пл;Техники;СХ_Культ;200СХ;Обст;Спас_Л;Спас_Т;Спас_Ск;Направление;Подпись", ";")
    
    For i = 0 To UBound(markers)
        marker = markers(i)
        SetData wrd, marker, defaultVal
    Next i
End Sub





'-------------------------Функции формирования сложных строк------------------------
'Public Function GetTechniks(ByRef gfsShapes As Collection, ByRef units As Collection) As String        'На будущее
Public Function GetTechniks(ByRef units As Collection) As String
'Возвращает свормированнуб строку "Тип, количество и принадлежность пожарной техники"
Dim unit As Variant
Dim shpColl As Collection
'Dim shp As Visio.Shape
Dim tmpStr As String
Dim mainStr As String

    mainStr = ""
    For Each unit In units
        tmpStr = ""
        'АЦ
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipAC & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "АЦ:", ", ")
        'АЛ
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipAL & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "АЛ:", ", ")
        'АНР
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipANR & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "АНР:", ", ")
        'АШ
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipASH & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "АШ:", ", ")
        'АСО
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipASO & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "АСО:", ", ")
        'КС
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipKS & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "КС:", ", ")
        'Дальше добавить отслаьные типа автомобилей согласно их IndexPers
        
        'В конце добавляем название подразделения (если хоть одна единица техники была найдена)
        If tmpStr <> "" Then mainStr = mainStr & Left(tmpStr, Len(tmpStr) - 2) & "(" & unit & "); "
    Next unit
    
GetTechniks = mainStr
End Function

Private Function GetVictims() As Collection
Dim shp As Visio.Shape
Dim col As Collection
Dim deadCount As Integer
Dim casCount As Integer
Dim deads As String
Dim cased As String
Dim i As Integer
    
    Set col = A.GetGFSShapes("User.IndexPers:" & indexPers.ipPostradavshie)
    deads = " "
    cased = " "
    
    For Each shp In col
        deadCount = cellVal(shp, "Prop.CasCount")
        casCount = cellVal(shp, "Prop.iedCount")
        
        For i = 1 To 5
            If deadCount + casCount > 5 Then Exit For
            
            If i <= deadCount Then
                deads = deads & cellVal(shp, "Prop.Cas" & i, visUnitsString) & ", "
            Else
                cased = cased & cellVal(shp, "Prop.Cas" & i, visUnitsString) & ", "
            End If
        Next i
    Next shp
    
    Set GetVictims = New Collection
    GetVictims.Add deads
    GetVictims.Add cased
End Function




Private Function GetMarkStr(ByRef col As Collection) As String
'Формируем коллекцию оценок
Dim i As Integer
Dim rowName As String
Dim marks As String
Dim shp As Visio.Shape
    
    On Error GoTo ex
    
    For Each shp In col
        For i = 0 To shp.RowCount(visSectionUser) - 1
            rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
            If Len(rowName) > 9 Then
                If Left(rowName, 9) = "GFS_Info_" Then
                    'Печатаем в поле оценки
                    If IsGFSShapeWithIP(shp, ipDutyFace, True) Then
                        marks = marks & cellVal(shp, "Prop.Duty", visUnitsString) & " " & GetInfo(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString)) & Chr(13)
                    Else
                        marks = marks & GetInfo(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString)) & ", "
                    End If
                    
'                    cmndID = cmndID + 1
                End If
            End If
        Next i
    Next shp
    
    GetMarkStr = marks

Exit Function
ex:
    GetMarkStr = " "
End Function

Private Function GetInfo(ByVal str As String) As String
Dim strArr() As String
    
    strArr = Split(str, delimiter)
    If UBound(strArr) > 1 Then
        GetInfo = strArr(UBound(strArr))
    ElseIf UBound(strArr) = 1 Then
        GetInfo = strArr(0)
    Else
        GetInfo = " "
    End If
End Function

Private Function GetStabMembers() As String
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim gettedCol As Collection
Dim stabMember As String
    
    On Error GoTo ex
    
    Set gettedCol = A.GetGFSShapes("User.IndexPers:66")
    
    If gettedCol.Count > 0 Then
        Set shp1 = gettedCol(1)
        
        Set gettedCol = A.GetGFSShapes("User.IndexPers:65")
        For Each shp2 In gettedCol
            If shp2.SpatialRelation(shp1, 0, 0) = 4 Then
                stabMember = stabMember & cellVal(shp2, "Prop.Duty", visUnitsString) & ": " & _
                    cellVal(shp2, "Prop.FIO", visUnitsString) & ", "
            End If
        Next shp2
    End If
    
GetStabMembers = stabMember
Exit Function
ex:
    GetStabMembers = " "
End Function

Private Function GetBUSTPString() As String
Dim gettedCol As Collection
Dim shp As Visio.Shape
Dim tmpStr As String

    On Error GoTo ex

    Set gettedCol = A.GetGFSShapes("Prop.UTP_STP_Reserv:Участок;Prop.UTP_STP_Reserv:Сектор")
    
    For Each shp In gettedCol
        tmpStr = tmpStr & cellVal(shp, "User.IndexPers.Prompt", visUnitsString) & ": " & _
            "Начальник - " & cellVal(shp, "Prop.NachUTP", visUnitsString) & ", " & _
            "задача - " & cellVal(shp, "Prop.UTPMission", visUnitsString) & ", " & _
            "приданые СиС - " & cellVal(shp, "Prop.UTPUnits", visUnitsString) & Chr(13)
    Next shp
    
GetBUSTPString = tmpStr
Exit Function
ex:
    GetBUSTPString = " "
End Function

Private Function GetServicesCommunications() As String
Dim gettedCol As Collection
Dim shp As Visio.Shape
Dim tmpStr As String
    
    On Error GoTo ex
    
    Set gettedCol = A.GetGFSShapes("Prop.ServiceMembership:Прочие")
    
    For Each shp In gettedCol
        tmpStr = tmpStr & cellVal(shp, "Prop.ServiceDescription", visUnitsString) & ", "
    Next shp
    
    GetServicesCommunications = tmpStr
    
Exit Function
ex:
    GetServicesCommunications = " "
End Function




'Public Function TTT()
'Dim c As Collection
'
'    Set c = New Collection
'    c.Add "ПСЧ-20"
'    c.Add "ПСЧ-10"
'    c.Add "ПЧ-13"
'    c.Add "П-1"
'    Debug.Print GetTechniks(c)
''    Debug.Print A.Refresh(1).GetGFSShapesAnd("User.IndexPers:" & indexpers.ipAC & ";Prop.Unit:ПСЧ-20")
'End Function



