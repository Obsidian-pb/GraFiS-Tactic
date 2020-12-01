VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_InsertAnalizedCell 
   Caption         =   "Выбор вычисляемой ячейки"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   OleObjectBlob   =   "f_InsertAnalizedCell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_InsertAnalizedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public formShape As Visio.Shape



Private Sub UserForm_Activate()
    FillAElemList
End Sub

Public Sub ShowMe(ByRef shp As Visio.Shape)
    Set formShape = shp
    
    Me.Show
End Sub










Private Sub cb_Cancel_Click()
    Me.Hide
End Sub
Private Sub cb_Save_Click()
    SaveCells
'    Me.Hide
End Sub



Private Sub SaveCells()
Dim i As Integer
Dim rowI As Integer
    
    For i = 0 To Me.lb_AElements.ListCount - 1
        If Me.lb_AElements.Selected(i) Then
            If formShape.CellExists("Prop." & Me.lb_AElements.List(i, 0), 0) = 0 Then
                rowI = formShape.AddNamedRow(visSectionProp, Me.lb_AElements.List(i, 0), visTagDefault)
'                formShape.CellsSRC(visSectionProp, rowI, visCustPropsLabel).FormulaU = """" & Me.lb_AElements.List(i, 1) & """"
            End If
        End If
    Next i
    
End Sub


Private Sub FillAElemList()
Dim i As Integer
Dim arr As Variant
    
    arr = a.GetElementsCode
    
    Me.lb_AElements.Clear
    For i = 0 To UBound(arr, 1)
        Me.lb_AElements.AddItem "A_" & arr(i, 0), i
        Me.lb_AElements.List(i, 1) = arr(i, 1)
    Next i

End Sub
