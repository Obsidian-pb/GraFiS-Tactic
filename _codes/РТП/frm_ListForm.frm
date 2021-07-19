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


Public Function Activate(ByVal arr As Variant, Optional ByVal colWidth As String = "")
Dim colCount As Byte
    
    colCount = UBound(arr, 1) + 1
    Me.LB_List.ColumnCount = colCount
    
    If colWidth <> "" Then
        Me.LB_List.ColumnWidths = colWidth
    End If
    
    
    Me.LB_List.List = arr
    
    Me.Show
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
