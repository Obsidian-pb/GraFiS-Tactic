VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListForm 
   Caption         =   "Список"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11040
   OleObjectBlob   =   "frm_ListForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ListForm"
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

Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Public WithEvents menuButtonExportToWord As CommandBarButton
Attribute menuButtonExportToWord.VB_VarHelpID = -1



Private Sub Stretch()
'Устанавливаем размер содержимого окна
    Me.LB_List.width = Me.width - con_BorderWidth
    Me.LB_List.height = Me.height - con_BorderHeightForList
End Sub



Private Sub UserForm_Resize()
    Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    wAddon.Close
End Sub


Public Function Activate(ByVal arr As Variant, ByVal colWidth As String, _
                         ByVal frmID As String, ByVal frmCaption As String) As frm_ListForm
Dim colCount As Byte
    
    
    'Наполняем содержимым список
    colCount = UBound(arr, 2) + 1
    Me.LB_List.ColumnCount = colCount
    
    If colWidth <> "" Then
        Me.LB_List.ColumnWidths = colWidth
    End If
    
    Me.LB_List.List = arr
    
    'Показываем привязанную к приложению форму
    Set wAddon = ActiveWindow.Windows.Add(frmID, visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 300, 300, "LST", "LST")
    
    Me.Caption = frmID
    FormHandle = FindWindow(vbNullString, frmID)
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = frmCaption
    
Me.Show
Set Activate = Me
End Function




Private Sub LB_List_Change()
Dim shpID As Long
    
    On Error GoTo EX
    
    'Выделояем фигуру
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    Application.ActiveWindow.Select Application.ActivePage.Shapes.ItemFromID(shpID), visDeselectAll + visSelect
    
EX:
End Sub

Private Sub LB_List_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim shpID As Long
Dim shp As Visio.Shape

    On Error GoTo EX
    
    'Определяем фигуру для которой сделана запись
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    Set shp = Application.ActivePage.Shapes.ItemFromID(shpID)
    Application.ActiveWindow.Select shp, visDeselectAll + visSelect
    
    'Устанавливаем фокус на фигуре
    Application.ActiveWindow.Zoom = 1.5 * GetScaleAt200
    Application.ActiveWindow.ScrollViewTo shp.Cells("PinX"), shp.Cells("PinY")
EX:
End Sub







'------------------Работа с всплывающим меню------------------
Private Sub LB_List_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        CreateNewMenu
    End If
End Sub


Private Sub CreateNewMenu()
'Создаём всплывающее меню мастера проверок
Dim popupMenuBar As CommandBar
Dim Ctrl As CommandBarControl
    
    'Получаем ссылку на всплывающее меню
    GetToolBar popupMenuBar, "ContextListMenu", msoBarPopup
    
    'Очищаем имеющиеся пункты меню
    For Each Ctrl In popupMenuBar.Controls
        Ctrl.Delete
    Next
    
    'Добавляем новые кнопки
    Set menuButtonExportToWord = NewPopupItem(popupMenuBar, 1, 214, "Экспортировать в Word")
'    Set menuButtonRestore = NewPopupItem(popupMenuBar, 1, 213, "Показать все скрытые замечания" & " (" & remarksHided & ")", , remarksHided <> 0)
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

'---------------------Экспорт в Ворд---------------------------------
Public Sub ExportToWord()
'Экспортируем содержимое списка в документ Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
Dim wrdTblRow As Object

Dim i As Integer
Dim j As Integer
Dim colCount As Byte

    
    
    colCount = Me.LB_List.ColumnCount - 1
    
    'Создаем новый документ Word
    Set wrd = CreateObject("Word.Application")
    wrd.visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add
    wrdDoc.Activate
    'Создаем в новом документе таблицу требуемых размеров
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, Me.LB_List.ListCount, colCount)
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
    
    
    'Заполняем таблицу
    For i = 0 To Me.LB_List.ListCount - 1
        For j = 1 To colCount
            wrdTbl.Rows(i + 1).Cells(j).Range.text = Me.LB_List.Column(j, i)
        Next j
    Next i
    
    wrdTbl.AutoFitBehavior 1        'Устанавливаем ширину столбцов по содержимому
End Sub



