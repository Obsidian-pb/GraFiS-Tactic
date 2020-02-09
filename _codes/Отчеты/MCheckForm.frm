VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MCheckForm 
   Caption         =   " Мастер проверок схемы - Бета-версия"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8625
   OleObjectBlob   =   "MCheckForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MCheckForm"
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

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private f_MCheckForm As MCheckForm
Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Private Const con_BorderWidth = 6
Private Const con_BorderHeightForList = 20

Public WithEvents menuButtonHide As CommandBarButton
Attribute menuButtonHide.VB_VarHelpID = -1
Public WithEvents menuButtonRestore As CommandBarButton
Attribute menuButtonRestore.VB_VarHelpID = -1
Public WithEvents menuButtonOptions As CommandBarButton
Attribute menuButtonOptions.VB_VarHelpID = -1


'Private Sub ListBox1_Click()
'    MasterCheckRefresh
'End Sub


 
Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal x As Single, ByVal y As Single)
    If Button = 2 Then
        If y > 0 And x > 0 And y < ListBox1.Height And x < ListBox1.Width Then
            DoEvents
            CreateNewMenu
        End If
    End If
 End Sub


Private Sub UserForm_Activate()
    Set wAddon = ActiveWindow.Windows.Add("MCheck", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)

    Me.Caption = "MCheck"
    FormHandle = FindWindow(vbNullString, "MCheck")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "Мастер проверок"
    

    
    Set app = Visio.Application
End Sub


Private Sub pS_Stretch()
'Устанавливаем размер содержимого окна
    Me.MultiPage1.Width = Me.InsideWidth - con_BorderWidth
    Me.MultiPage1.Height = Me.InsideHeight - con_BorderWidth
    Me.ListBox1.Width = Me.MultiPage1.Width - con_BorderWidth
    Me.ListBox1.Height = Me.MultiPage1.Height - con_BorderHeightForList
    Me.ListBox2.Width = Me.MultiPage1.Width - con_BorderWidth
    Me.ListBox2.Height = Me.MultiPage1.Height - con_BorderHeightForList
End Sub

Private Sub UserForm_Deactivate()
    Set app = Nothing
End Sub

Private Sub UserForm_Resize()
    pS_Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    wAddon.Close
End Sub

Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
   MasterCheckRefresh
End Sub


'------------------Работа с всплывающим меню------------------
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

Private Sub CreateNewMenu()
'Создаём всплывающее меню мастера проверок
Dim popupMenuBar As CommandBar
Dim ctrl As CommandBarControl
    
    'Получаем ссылку на всплывающее меню
    GetToolBar popupMenuBar, "ContextMenuListBox", msoBarPopup
    
    'Очищаем имеющиеся пункты меню
    For Each ctrl In popupMenuBar.Controls
        ctrl.Delete
    Next
    
    'Добавляем новые кнопки
    Set menuButtonHide = NewPopupItem(popupMenuBar, 1, 214, "Не учитывать выделенное замечание")
    Set menuButtonRestore = NewPopupItem(popupMenuBar, 1, 213, "Показать все скрытые замечания" & " (" & nX & ")", , nX <> 0)
    Set menuButtonOptions = NewPopupItem(popupMenuBar, 1, 212, "Настроить выделенное замечание", , False)
    
    'Показываем меню
    popupMenuBar.ShowPopup
End Sub

Private Sub GetToolBar(ByRef toolBar As CommandBar, ByVal toolBarName As String, ByVal barPosition As MsoBarPosition)
    On Error Resume Next
    'Пытаемся получить ссылку на всплывающее меню
    Set toolBar = Application.CommandBars(toolBarName)

    'Если такого меню нет, создаем его
    If toolBar Is Nothing Then
        Set toolBar = Application.CommandBars.Add(toolBarName, barPosition)
    End If
    
End Sub

Private Sub menuButtonHide_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    HideComment
End Sub

Private Sub menuButtonRestore_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    RestoreComment
End Sub


