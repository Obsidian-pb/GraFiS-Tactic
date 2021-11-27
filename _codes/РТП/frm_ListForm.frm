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

'Private curShapeID As Long
'Private lastCalcTime As Double
'Const recalcInterval = 0.0000116

'Private elemCollection As Collection

'--------------------------Основные процедуры и функции класса--------------------
'Public Function Activate() As TacticDataForm
'    'Потом переписать по человечески! Если другие формы анализа не показаны, то для новой указывается высота, иначе - нет. Нужно чтоб корректно отображались формы
'    If WarningsForm.visible = True Then
'        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom + visWSAnchorMerged, visAnchorBarAddon, , , 600, , "MC", "MC")
'    Else
'        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 600, 210, "MC", "MC")
'    End If
'
'    Me.Caption = "TacticDataForm"
'    FormHandle = FindWindow(vbNullString, "TacticDataForm")
'    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
'    SetParent FormHandle, wAddon.WindowHandle32
'    wAddon.Caption = "Мастер тактических данных"
'
'    'Активируем экземпляр объекта приложения для отслеживания изменений ячеек
'    Set app = Visio.Application
'
'    'Показываем форму
'    Me.Show
'
'Set Activate = Me
'End Function

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










Public Function Activate(ByVal arr As Variant, Optional ByVal colWidth As String = "") As frm_ListForm
Dim colCount As Byte
    
    
    'Наполняем содержимым список
    colCount = UBound(arr, 1) + 1
    Me.LB_List.ColumnCount = colCount
    
    If colWidth <> "" Then
        Me.LB_List.ColumnWidths = colWidth
    End If
    
    Me.LB_List.List = arr
    
    'Показываем привязанную к приложению форму
    Set wAddon = ActiveWindow.Windows.Add("Lists", visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 300, 300, "LST", "LST")
    
    Me.Caption = "Lists"
    FormHandle = FindWindow(vbNullString, "Lists")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "Списки"
    
Me.Show
Set Activate = Me
End Function






Private Sub LB_List_Change()
Dim shpID As Long
    
    On Error GoTo ex
    
    'Выделояем фигуру
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    Application.ActiveWindow.Select Application.ActivePage.Shapes.ItemFromID(shpID), visDeselectAll + visSelect
    
ex:
End Sub




Private Sub LB_List_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim shpID As Long
Dim shp As Visio.Shape

    On Error GoTo ex
    
    'Определяем фигуру для которой сделана запись
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    Set shp = Application.ActivePage.Shapes.ItemFromID(shpID)
    Application.ActiveWindow.Select shp, visDeselectAll + visSelect
    
    'Устанавливаем фокус на фигуре
    Application.ActiveWindow.Zoom = 1.5 * GetScaleAt200
    Application.ActiveWindow.ScrollViewTo shp.Cells("PinX"), shp.Cells("PinY")
ex:
End Sub
