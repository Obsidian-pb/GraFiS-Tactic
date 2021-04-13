VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_LinkToCell 
   Caption         =   "Выберите связываемую ячейку"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10440
   OleObjectBlob   =   "f_LinkToCell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_LinkToCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public formShape As Visio.Shape
Public frmlShapesColl As Collection





Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    SaveLink
End Sub

Private Sub lb_FormulaShapes_Change()
    If Me.lb_FormulaShapes.ListIndex < 0 Then Exit Sub
    FillFormulaList frmlShapesColl(Me.lb_FormulaShapes.ListIndex + 1)
End Sub



Private Sub lb_FormulaShapes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim sourceShp As Visio.Shape
    
    Set sourceShp = frmlShapesColl(Me.lb_FormulaShapes.ListIndex + 1)
    Application.ActiveWindow.ScrollViewTo sourceShp.Cells("PinX"), sourceShp.Cells("PinY")

End Sub

Private Sub UserForm_Activate()
    FillShapesList
End Sub



Public Sub ShowMe(ByRef shp As Visio.Shape)
    Set formShape = shp
    
    Me.Show
End Sub







Private Sub FillShapesList()
Dim shp As Visio.Shape
Dim frmlName As String
    
    Set frmlShapesColl = New Collection
    Me.lb_FormulaShapes.Clear
    For Each shp In Application.ActivePage.Shapes
        If IsGFSShapeWithIP(shp, 500) Then
            frmlName = cellVal(shp, "Prop.Название", visUnitsString)
            AddUniqueCollectionItem frmlShapesColl, shp
            If frmlName = "0" Then
                Me.lb_FormulaShapes.AddItem shp.name
            Else
                Me.lb_FormulaShapes.AddItem frmlName
            End If
        End If
    Next shp
End Sub

Private Sub FillFormulaList(ByRef shp As Visio.Shape)
Dim i As Integer
Dim cell As Visio.cell
    
    Me.lb_Cells.Clear
    For i = 0 To shp.RowCount(visSectionProp) - 1
        Me.lb_Cells.AddItem shp.CellsSRC(visSectionProp, i, visTagDefault).RowNameU
    Next i
    
End Sub


Private Sub SaveLink()
Dim i As Integer
Dim rowI As Integer
Dim rowName As String
Dim sourceShp As Visio.Shape
Dim frml As String
    
    Set sourceShp = frmlShapesColl(Me.lb_FormulaShapes.ListIndex + 1)
    For i = 0 To Me.lb_Cells.ListCount - 1
        If Me.lb_Cells.Selected(i) Then
            rowName = Me.lb_Cells.List(i, 0) & "_" & sourceShp.ID
            If Not ShapeHaveCell(formShape, "Prop." & rowName) Then
                rowI = formShape.AddNamedRow(visSectionProp, rowName, visTagDefault)
                frml = "Sheet." & sourceShp.ID & "!Prop." & Me.lb_Cells.List(i, 0)
                formShape.CellsSRC(visSectionProp, rowI, visCustPropsValue).FormulaU = frml
            End If
        End If
    Next i
End Sub
