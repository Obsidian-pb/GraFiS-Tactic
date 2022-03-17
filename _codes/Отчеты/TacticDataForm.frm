VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TacticDataForm 
   Caption         =   "Тактические данные"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   OleObjectBlob   =   "TacticDataForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TacticDataForm"
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

'Public WithEvents menuButtonHide As CommandBarButton
'Public WithEvents menuButtonRestore As CommandBarButton
'Public WithEvents menuButtonOptions As CommandBarButton
Public WithEvents menuButtonExportToWord As CommandBarButton
Attribute menuButtonExportToWord.VB_VarHelpID = -1

Private curShapeID As Long
Private lastCalcTime As Double
Const recalcInterval = 0.0000116

'Private elemCollection As Collection

'--------------------------Основные процедуры и функции класса--------------------
Public Function Activate() As TacticDataForm
    'Потом переписать по человечески! Если другие формы анализа не показаны, то для новой указывается высота, иначе - нет. Нужно чтоб корректно отображались формы
    If WarningsForm.Visible = True Then
        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom + visWSAnchorMerged, visAnchorBarAddon, , , 600, , "MC", "MC")
    Else
        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 600, 210, "MC", "MC")
    End If
    
    Me.Caption = "TacticDataForm"
    FormHandle = FindWindow(vbNullString, "TacticDataForm")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "Мастер тактических данных"
        
    'Активируем экземпляр объекта приложения для отслеживания изменений ячеек
    Set app = Visio.Application
    
    'Показываем форму
    Me.Show
    
Set Activate = Me
End Function

Private Sub Stretch()
'Устанавливаем размер содержимого окна
    Me.lstTacticData.Width = Me.Width - con_BorderWidth
    Me.lstTacticData.Height = Me.Height - con_BorderHeightForList
End Sub

Private Sub UserForm_Resize()
    Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    Set app = Nothing
    wAddon.Close
    
''---Скрываем кнопку "Экспорт в Word"
'    Application.CommandBars("Спецфункции").Controls("Экспорт в Word").Visible = False
End Sub

Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
    
    If curShapeID <> Cell.Shape.ID Then
        Refresh
        curShapeID = Cell.Shape.ID
        'Debug.Print Cell.Shape.Name
    Else
        If lastCalcTime + recalcInterval < Now() Then
            Refresh
            'Debug.Print Cell.Shape.Name
        End If
    End If
    
    lastCalcTime = Now()
End Sub



'------------Процедуры обновления формы--------------------------
Public Sub Refresh()
'Обновляем содержимое списка предупреждений
Dim i As Integer
Dim elemCollection As Collection
Dim elem As Element

'---Проводим расчет элементов
    A.Refresh Application.ActivePage.Index
    
'---Очищаем форму и задаем условия по-умолчанию
    Me.lstTacticData.Clear

'---Запускаем условия обработки
    Set elemCollection = A.elements.GetElementsCollection("")
    
    i = 0
'    With A
    For Each elem In elemCollection
        If elem.inTacticForm Then
            If elem.Result > 0 And elem.Result <> "" Then
                lstTacticData.AddItem elem.callName, i
                lstTacticData.List(i, 1) = elem.ResultStr
                i = i + 1
            End If
        End If
    Next elem
'    End With
    
    'Добавляем в конце пустую строку, для корректного отображения больших списков
    lstTacticData.AddItem " "

End Sub

'------------Процедуры экспорта данных формы--------------------------
Public Sub ExportToWord()
'Экспортируем данные формы в документ Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
Dim tactDataRowsCount As Integer
Dim i As Integer
    
    tactDataRowsCount = lstTacticData.ListCount
    
    'Создаем новый документ Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add
    wrdDoc.Activate
    'Создаем в новом документе таблицу требуемых размеров
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, tactDataRowsCount, 2)
    With wrdTbl
        If .style <> "Сетка таблицы" Then
            .style = "Сетка таблицы"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    
    'Заполняем табблицу тактическими данными
    For i = 1 To tactDataRowsCount
        If Not IsNull(lstTacticData.List(i - 1, 1)) Then
            wrdTbl.Rows(i).Cells(1).Range.Text = lstTacticData.List(i - 1, 0)
            wrdTbl.Rows(i).Cells(2).Range.Text = lstTacticData.List(i - 1, 1)
        End If
    Next i
End Sub




Private Sub UserForm_Terminate()
''---Скрываем кнопку "Экспорт в Word"
'    Application.CommandBars("Спецфункции").Controls("Экспорт в Word").Visible = False
End Sub





'------------------Работа с всплывающим меню------------------
Private Sub lstTacticData_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        CreateNewMenu
    End If
End Sub


Private Sub CreateNewMenu()
'Создаём всплывающее меню мастера проверок
Dim popupMenuBar As CommandBar
Dim Ctrl As CommandBarControl
    
    'Получаем ссылку на всплывающее меню
    GetToolBar popupMenuBar, "ContextListMenuTacic", msoBarPopup
    
    'Очищаем имеющиеся пункты меню
    For Each Ctrl In popupMenuBar.Controls
        Ctrl.Delete
    Next
    
    'Добавляем новые кнопки
    Set menuButtonExportToWord = NewPopupItem(popupMenuBar, 1, 268, "Экспортировать в Word")
    
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
'        .beginGroup = beginGroup
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

Private Sub menuButtonExportToWord_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ExportToWord
End Sub

