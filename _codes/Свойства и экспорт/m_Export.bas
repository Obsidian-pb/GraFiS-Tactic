Attribute VB_Name = "m_Export"
Option Explicit

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
        SetData wrd, "НП", cellval(gfsShapes, "Prop.NP_Name", visUnitsString, "")                       'Населенный пункт
        SetData wrd, "ДЗФИО", cellval(gfsShapes, "Prop.PersonCreate", visUnitsString, "")                      'должность, звание, фамилия, имя, отчество (при наличии)
        SetData wrd, "Наименование", cellval(gfsShapes, "Prop.ObjectName", visUnitsString, "")          'Наименование
        SetData wrd, "Принадлежность", cellval(gfsShapes, "Prop.Affiliation", visUnitsString, "")       'Принадлежность объекта
        SetData wrd, "Адрес", cellval(gfsShapes, "Prop.Address", visUnitsString, "")                    'Адрес организации
        SetData wrd, "МестоПожара", cellval(gfsShapes, "Prop.FireStartPlace", visUnitsString, "")              'Место возникновения пожара
        SetData wrd, "Заявитель", cellval(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString, "")               'Фамилия, имя, отчество (при наличии) лица, обнаружившего пожар и способ сообщения о нем в пожарную охрану
        SetData wrd, "Заявитель_т", cellval(gfsShapes, "Prop.CallPhone", visUnitsString, "")            'Номер телефона заявителя
        
        
        
        
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
            SetData wrd, "Обст_Приб", cellval(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
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
        gettedTxt = cellval(gfsShapes, "Prop.HumansDie", visUnitsString, "")
'        SetData wrd, "200", "Погибло людей: " & Split(gettedTxt, "/")(0) & "в том числе детей: " & Split(gettedTxt, "/")(1) & "работников ПО: " & Split(gettedTxt, "/")(2)
        SetData wrd, "200", Split(gettedTxt, "/")(0)
        SetData wrd, "200Д", Split(gettedTxt, "/")(1)
        SetData wrd, "200ПО", Split(gettedTxt, "/")(2)
        
        
        
         'Травмировано людей
        gettedTxt = cellval(gfsShapes, "Prop.HumansInjured", visUnitsString, "")
        SetData wrd, "300", Split(gettedTxt, "/")(0)
        SetData wrd, "300Д", Split(gettedTxt, "/")(1)
        SetData wrd, "300ПО", Split(gettedTxt, "/")(2)
        
        'Информация о погибших и травмированных
        Set gettedCol = GetVictims
        SetData wrd, "200Св", gettedCol(1)
        SetData wrd, "300Св", gettedCol(2)
        
        'Уничтожено/повреждено
        '---строений
        gettedTxt = cellval(gfsShapes, "Prop.ConstructionsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Стр", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Стр", Split(gettedTxt, "/")(1)
        '---Квартир
        gettedTxt = cellval(gfsShapes, "Prop.FlatsAffected", visUnitsString, "")
        SetData wrd, "Уничт_ЖК", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_ЖК", Split(gettedTxt, "/")(1)
        '---Комнат
        gettedTxt = cellval(gfsShapes, "Prop.RoomsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Комн", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Комн", Split(gettedTxt, "/")(1)
        '---Площади
        gettedTxt = cellval(gfsShapes, "Prop.SquareAffected", visUnitsString, "")
        SetData wrd, "Уничт_Пл", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Пл", Split(gettedTxt, "/")(1)
        '---Техники
        gettedTxt = cellval(gfsShapes, "Prop.TechnicsAffected", visUnitsString, "")
        SetData wrd, "Уничт_Тех", Split(gettedTxt, "/")(0)
        SetData wrd, "Повр_Тех", Split(gettedTxt, "/")(1)
        '---сельхоз культур
        gettedTxt = cellval(gfsShapes, "Prop.AgricultureAffected", visUnitsString, "")
        SetData wrd, "Уничт_СХ", ClearString(gettedTxt)
        '---сельхоз животных
        gettedTxt = cellval(gfsShapes, "Prop.CattleAffected", visUnitsString, "")
        SetData wrd, "200СХ", ClearString(gettedTxt)
        'Спасено
        '---людей
        gettedTxt = cellval(gfsShapes, "Prop.Saved", visUnitsString, "")
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
                     
        
        
        
        SetData wrd, "ПожАвт", cellval(gfsShapes, "Prop.FireAutomatics", visUnitsString, "")
        SetData wrd, "Осн_Спец", "Основных ПА: " & A.Result("MainOverallHave") & ", Специальных ПА:" & A.Result("SpecialPAHave")
        
        
        'Обстоятельства усложняющие обстановку
        gettedTxt = cellval(gfsShapes, "Prop.CircumstancesRize", visUnitsString, "")
        SetData wrd, "Обст", ClearString(gettedTxt)
        
        'Силы и средства применявшиеся при тушении
        gettedTxt = cellval(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "СиС", ClearString(gettedTxt)
        
        'Водоисточники
        gettedTxt = cellval(gfsShapes, "Prop.WaterSources", visUnitsString, "")
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
    '---Даты и время
        'Дата пожара:
        gettedDate = CDate(A.Result("FireTime"))
        SetData wrd, "П_Дата", Format(gettedDate, "DD.MM.YYYY")



'        gettedDate = CDate(A.Result("FireTime"))
'        SetData wrd, "П_День", Format(gettedDate, "DD")                             'День возникновения пожара
'        SetData wrd, "П_Месяц", Split(Format(gettedDate, "DD MMMM"), " ")(1)        'Месяц возникновения пожара
'        SetData wrd, "П_Год", Format(gettedDate, "YY")                              'Месяц возникновения пожара
'        'Дата сообщения:
'        gettedDate = CDate(A.Result("InfoTime"))
'        SetData wrd, "Сооб_Дата", Format(gettedDate, "DD MMMM YYYY")                'Дата сообщения о пожаре
'        SetData wrd, "Сооб_Час", Format(gettedDate, "HH")                           'Час сообщения о пожаре
'        SetData wrd, "Сооб_Мин", Format(gettedDate, "NN")                           'Минута сообщения о пожаре
'        'Время прибытия первого подразделения:
'        gettedDate = CDate(A.Result("FirstArrivalTime"))
'        SetData wrd, "1Подр_Час", Format(gettedDate, "HH")                           'Час сообщения о пожаре
'        SetData wrd, "1Подр_Мин", Format(gettedDate, "NN")                           'Минута сообщения о пожаре
'
'
'        'Перечень подразделений
'        Set gettedCol = GetUniqueVals(gfsShapes, "Prop.Unit", , "-", "-")
'        SetData wrd, "Подразделения", StrColToStr(gettedCol, ", ")
'        'Тип, количество и принадлежность пожарной техники
'        SetData wrd, "Техника", GetTechniks(gettedCol)
        
        
        
        

        
        
End Sub
'Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String, _
'                    Optional ByVal ignore As Variant = 0, Optional ByVal ifIgnore As Variant = " ")
Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String)
'Замена маркера в тексте текстом из результата анализа
    
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
        deadCount = cellval(shp, "Prop.CasCount")
        casCount = cellval(shp, "Prop.iedCount")
        
        For i = 1 To 5
            If deadCount + casCount > 5 Then Exit For
            
            If i <= deadCount Then
                deads = deads & cellval(shp, "Prop.Cas" & i, visUnitsString) & ", "
            Else
                cased = cased & cellval(shp, "Prop.Cas" & i, visUnitsString) & ", "
            End If
        Next i
    Next shp
    
    Set GetVictims = New Collection
    GetVictims.Add deads
    GetVictims.Add cased
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



