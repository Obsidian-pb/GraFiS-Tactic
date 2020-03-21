VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WarningsForm 
   Caption         =   "Предупреждения"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   OleObjectBlob   =   "WarningsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WarningsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Win64 Then
    #If VBA7 Then
        Public FormHandle As LongPtr
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As LongPtr, _
                        ByVal nIndex As LongPtr, _
                        ByVal dwNewLong As Long) As LongPtr
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As LongPtr, _
                        ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetParent Lib "user32" ( _
                        ByVal hWndChild As LongPtr, _
                        ByVal hWndNewParent As LongPtr) As LongPtr
    #Else
        Public FormHandle As Long
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #End If
#Else
    #If VBA7 Then
        Public FormHandle As Long
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #Else
        Public FormHandle As Long
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #End If
#End If


Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Private Const con_BorderWidth = 6
Private Const con_BorderHeightForList = 6

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Private remarks() As String         'Массив сведений для отображения
Const remarksItems = 28             'Верхний индекс предупреждений (от 0, т.е. количество равно remarksItems+1)
Private remarksHided As Integer     'Переменная количества скрытых замечаний

Public WithEvents menuButtonHide As CommandBarButton
Attribute menuButtonHide.VB_VarHelpID = -1
Public WithEvents menuButtonRestore As CommandBarButton
Attribute menuButtonRestore.VB_VarHelpID = -1
'Public WithEvents menuButtonOptions As CommandBarButton




'--------------------------Основные процедуры и функции класса--------------------


Public Function Activate() As WarningsForm
    Set wAddon = ActiveWindow.Windows.Add("WarningsForm", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)

    Me.Caption = "WarningsForm"
    FormHandle = FindWindow(vbNullString, "WarningsForm")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "Мастер проверок"
        
    'Активируем экземпляр объекта приложения для отслеживания изменений ячеек
    Set app = Visio.Application
    
    'Показываем форму
    Me.Show
    
Set Activate = Me
End Function

Private Sub Stretch()
'Устанавливаем размер содержимого окна
    Me.lstWarnings.Width = Me.Width - con_BorderWidth
    Me.lstWarnings.Height = Me.Height - con_BorderHeightForList
End Sub

Private Sub UserForm_Initialize()
    ReDim remarks(remarksItems, 1)
End Sub

'Private Sub UserForm_Terminate()
'    Set hidedRemarks = Nothing
'End Sub

Private Sub UserForm_Resize()
    Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    Set app = Nothing
    wAddon.Close
End Sub

Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
    Refresh
End Sub

'------------Список предупреждений------------
Private Sub lstWarnings_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        If Y > 0 And X > 0 And Y < lstWarnings.Height And X < lstWarnings.Width Then
            DoEvents
            CreateNewMenu
        End If
    End If
End Sub




'------------Процедуры обновления формы--------------------------
Public Sub Refresh()
'Обновляем содержимое списка предупреждений
Dim i As Integer

    On Error GoTo EX
'---Проводим расчет элементов
    A.Refresh Application.ActivePage.Index
    
'---Очищаем форму и задаем условия по-умолчанию
    Me.lstWarnings.Clear
    
'---Запускаем условия обработки
    i = 0
    With A
        'Очаг
        If remarks(i, 1) = "" Then
            If .Result("OchagCount") = 0 And (.Result("SmokeCount") > 0 Or .Result("SpreadCount") > 0 Or .Result("FireCount") > 0) Then
                remarks(i, 0) = "Не указан очаг пожара"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SmokeCount") = 0 Then
                remarks(i, 0) = "Не указаны зоны задымления"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SpreadCount") = 0 Then
                remarks(i, 0) = "Не указаны пути распространения пожара"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        'Управление
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BUCount") >= 3 And .Result("ShtabCount") = 0 Then
                remarks(i, 0) = "Не создан оперативный штаб"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("RNBDCount") = 0 And .Sum("OchagCount;FireCount") > 0 Then
                remarks(i, 0) = "Не указано решающее направление"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("RNBDCount") > 1 Then
                remarks(i, 0) = "Решающее напраление должно быть одним"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BUCount") >= 5 And .Result("SPRCount") <= 1 Then
                remarks(i, 0) = "Не организованы секторы проведения работ"
            Else
                remarks(i, 0) = ""
            End If
        End If

        'ГДЗС
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSPBCount") < .Result("GDZSChainsCountWork") Then
                remarks(i, 0) = "Не выставлены посты безопасности для каждого звена ГДЗС (" & .Result("GDZSPBCount") & "/" & .Result("GDZSChainsCountWork") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsCountWork") >= 3 And .Result("GDZSKPPCount") Then
                remarks(i, 0) = "Не создан контрольно-пропускной пункт ГДЗС"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSDiscr") = True Then
                remarks(i, 0) = "В сложных условиях звенья ГДЗС должны состоять не менее чем из пяти газодымозащитников"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And Not .options.GDZSRezRoundUp Then
                remarks(i, 0) = "Недостаточно резервных звеньев ГДЗС с округлением в меньшую сторону (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And .options.GDZSRezRoundUp Then
                remarks(i, 0) = "Недостаточно резервных звеньев ГДЗС с округлением в большую сторону (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        'ППВ
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WaterSourceCount") > .Result("DistanceCount") Then
                remarks(i, 0) = "Не указаны расстояния от каждого водоисточника до места пожара (" & .Result("DistanceCount") & "/" & .Result("WaterSourceCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        'Рукава
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WorklinesCount") > .Result("LinesPosCount") Then
                remarks(i, 0) = "Не указаны положения (этаж) для каждой рабочей линии (" & .Result("LinesPosCount") & "/" & .Result("WorklinesCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("LinesCount") > .Result("LinesLableCount") Then
                remarks(i, 0) = "Не указаны подписи для каждой рукавной линии (" & .Result("LinesLableCount") & "/" & .Result("LinesCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        'План на местности
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BuildCount") > .Result("SOCount") Then
                remarks(i, 0) = "Не указаны подписи степени огнестойкости для каждого из зданий (" & .Result("SOCount") & "/" & .Result("BuildCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("OrientCount") = 0 And .Result("BuildCount") > 0 Then
                remarks(i, 0) = "Не указаны ориентиры на местности, такие как роза ветров или подпись улицы"
            Else
                remarks(i, 0) = ""
            End If
        End If

        'Показ расчетных данных
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("FactStreamW") <> 0 And .Result("FactStreamW") < .Result("NeedStreamW") Then
                remarks(i, 0) = "Недостаточный фактический расход воды (" & .Result("FactStreamW") & " л/c < " & .Result("NeedStreamW") & " л/с)"
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WaterValueNeed10min") > .Result("WaterValueHave") And PF_RoundUp(.Result("FactStreamW") / 32) > .Result("GetingWaterCount") Then
                remarks(i, 0) = "Недостаточный запас воды или Недостаточное водоснабжение боевых позиций"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("PersonnelHave") < .Result("PersonnelNeed") Then
                remarks(i, 0) = "Недостаточно личного состава, с учетом прибывшей техники (" & .Result("PersonnelHave") & "/" & .Result("PersonnelNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses51Have") < .Result("Hoses51Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 51 мм, с учетом прибывшей техники (" & .Result("Hoses51Have") & "/" & .Result("Hoses51Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses66Have") < .Result("Hoses66Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 66 мм, с учетом прибывшей техники (" & .Result("Hoses66Have") & "/" & .Result("Hoses66Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses77Have") < .Result("Hoses77Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 77 мм, с учетом прибывшей техники (" & .Result("Hoses77Have") & "/" & .Result("Hoses77Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses89Have") < .Result("Hoses89Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 89 мм, с учетом прибывшей техники (" & .Result("Hoses89Have") & "/" & .Result("Hoses89Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses110Have") < .Result("Hoses110Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 110 мм, с учетом прибывшей техники (" & .Result("Hoses110Have") & "/" & .Result("Hoses110Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses150Have") < .Result("Hoses150Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 150 мм, с учетом прибывшей техники (" & .Result("Hoses150Have") & "/" & .Result("Hoses150Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses200Have") < .Result("Hoses200Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 200 мм, с учетом прибывшей техники (" & .Result("Hoses200Have") & "/" & .Result("Hoses200Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses250Have") < .Result("Hoses250Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 250 мм, с учетом прибывшей техники (" & .Result("Hoses250Have") & "/" & .Result("Hoses250Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses300Have") < .Result("Hoses300Count") Then
                remarks(i, 0) = "Недостаточно напорных рукавов 300 мм, с учетом прибывшей техники (" & .Result("Hoses300Have") & "/" & .Result("Hoses300Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        '!!!ПРИ ДОБАВЛЕНИИ НОВЫХ ПОЗИЦИЙ ДЛЯ РАСЧЕТА НЕ ЗАБЫТЬ УВЕЛИЧИТЬ РАЗМЕР МАССИВА - remarksItems!!!
    End With
    
    
    'Формируем список предупреждений
    On Error Resume Next
        lstWarnings.List = GetWarningsListArray
    On Error GoTo EX
    
    'Если предупреждений не обнаружено, сообщаем об этом
    If lstWarnings.ListCount = 0 Then lstWarnings.AddItem "Замечаний не обнаружено"
    
    'Добавляем в конце пустую строку, для корректного отображения больших списков
    lstWarnings.AddItem " "

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "WarningForm.Refresh"
End Sub

Private Function GetWarningsListArray() As String()
'Возвращаем массив замечаний для заполнения lstWarnings
Dim i As Integer
Dim tmpArr() As String
Dim j As Integer
Dim size As Integer

    For i = 0 To UBound(remarks, 1)
        ' Абсолютно идиотский код, но в VBA, нельзя множество раз объявлять размер двухмерных массивов с сохранением данных, поэтому приходится сначала узнавать размер будущего массива, а потом сразу весь его объявлять((( злюсь
        If Not remarks(i, 0) = "" And remarks(i, 1) = "" Then
            size = size + 1
        End If
    Next i
    
    If size > 0 Then
        ReDim tmpArr(size - 1, 1)
        
        j = 0
        For i = 0 To UBound(remarks, 1)
            'Если имеется значение и при этом флаг скрытия не поставлен
            If Not remarks(i, 0) = "" And remarks(i, 1) = "" Then
                tmpArr(j, 0) = remarks(i, 0)
                tmpArr(j, 1) = i
                j = j + 1
            End If
        Next i
    End If
    
GetWarningsListArray = tmpArr
End Function

'---------Функции работы с настройками отображения комментариев
Private Sub RestoreComment()
'Обнуляем значения переменных не учитываемых замечений
Dim i As Integer
    
    For i = 0 To UBound(remarks, 1)
        remarks(i, 1) = ""
    Next
    
    remarksHided = 0
End Sub

Private Sub HideComment()
'Скрываем замечания по желанию пользователя
    If lstWarnings.Column(0, 0) = "Замечаний не обнаружено" Then Exit Sub
    
    If lstWarnings.ListIndex > -1 Then
        remarks(lstWarnings.Column(1, lstWarnings.ListIndex), 1) = "h"
        remarksHided = remarksHided + 1
    End If
End Sub


'------------------Работа с всплывающим меню------------------
Private Sub CreateNewMenu()
'Создаём всплывающее меню мастера проверок
Dim popupMenuBar As CommandBar
Dim Ctrl As CommandBarControl
    
    'Получаем ссылку на всплывающее меню
    GetToolBar popupMenuBar, "ContextMenuListBox", msoBarPopup
    
    'Очищаем имеющиеся пункты меню
    For Each Ctrl In popupMenuBar.Controls
        Ctrl.Delete
    Next
    
    'Добавляем новые кнопки
    Set menuButtonHide = NewPopupItem(popupMenuBar, 1, 214, "Не учитывать выделенное замечание")
    Set menuButtonRestore = NewPopupItem(popupMenuBar, 1, 213, "Показать все скрытые замечания" & " (" & remarksHided & ")", , remarksHided <> 0)
'    Set menuButtonOptions = NewPopupItem(popupMenuBar, 1, 212, "Опции замечаний")
    
    'Показываем меню
    popupMenuBar.ShowPopup
End Sub

Private Function NewPopupItem(ByRef commBar As CommandBar, ByVal itemType As Integer, ByVal itemFace As Integer, _
ByVal itemCaption As String, Optional ByVal beginGroup As Boolean = False, Optional ByVal enableTab As Boolean = True, _
Optional itemTag As String = "") As CommandBarControl
'Функция создает элемент контекстного меню и возвращает на него ссылку
Dim newControl As CommandBarControl

'    On Error Resume Next
    'Создаем новый контрол
    Set newControl = commBar.Controls.Add(itemType)
    
    'Указываем свойства нового контрола
    With newControl
        If itemFace > 0 Then .FaceID = itemFace
        .Tag = itemTag
        .Caption = itemCaption
        .beginGroup = beginGroup
        .Enabled = enableTab
    End With
    
Set NewPopupItem = newControl
End Function

Private Sub GetToolBar(ByRef toolBar As CommandBar, ByVal toolBarName As String, ByVal barPosition As MsoBarPosition)
    On Error Resume Next
    'Пытаемся получить ссылку на всплывающее меню
    Set toolBar = Application.CommandBars(toolBarName)

    'Если такого меню нет, создаем его
    If toolBar Is Nothing Then
        Set toolBar = Application.CommandBars.Add(toolBarName, barPosition)
    End If
    
End Sub

Private Sub menuButtonHide_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    HideComment
    Refresh
End Sub

Private Sub menuButtonRestore_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    RestoreComment
    Refresh
End Sub




