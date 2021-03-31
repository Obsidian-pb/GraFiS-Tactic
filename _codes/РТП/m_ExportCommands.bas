Attribute VB_Name = "m_ExportCommands"
Option Explicit

Private exl As Excel.Application
Private wkbk As Excel.Workbook
Private wkst As Excel.Worksheet
Private rowNumber As Integer


Public Sub Export()


Dim shp As Visio.Shape


    Set exl = New Excel.Application
    Set wkbk = exl.Workbooks.Add()
    Set wkst = exl.ActiveSheet
    exl.Visible = True
    
'    rowNumber = 1
'    For Each shp In Application.ActivePage.Shapes
'        fillCommand shp
'    Next shp
    
    rowNumber = 1
    For Each shp In Application.ActivePage.Shapes
        getSetTime shp
        
'        rowNumber = rowNumber + 1
    Next shp
    
End Sub

Private Sub fillCommand(ByRef shp As Visio.Shape)
Dim i As Integer
Dim rowName As String
Dim comTime As String
Dim comText As String
Dim comArr() As String
    
'    On Error GoTo ex
    
    For i = 0 To shp.RowCount(visSectionUser) - 1
        rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
        If Len(rowName) > 12 Then
            If Left(rowName, 12) = "GFS_Command_" Then
                comArr = Split(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), " | ")
                wkst.Cells(rowNumber, 1) = comArr(0)
                wkst.Cells(rowNumber, 2) = getCallName(shp) & " " & comArr(UBound(comArr))
                
                rowNumber = rowNumber + 1
            End If
        End If
        

    Next i
    
    
Exit Sub
ex:

End Sub

Public Function getCallName(ByRef shp As Visio.Shape) As String
    On Error GoTo ex
    getCallName = shp.Cells("Prop.Call").ResultStr(visUnitsString)
Exit Function
ex:
    getCallName = "-"
End Function
Public Sub getSetTime(ByRef shp As Visio.Shape)
    On Error GoTo ex
    If shp.Cells("User.IndexPers").Result(visNumber) = 34 Then
        wkst.Cells(rowNumber, 4) = shp.Cells("Prop.SetTime").ResultStr(visDate)
        wkst.Cells(rowNumber, 5) = shp.Cells("User.DiameterIn").ResultStr(visUnitsString)
        rowNumber = rowNumber + 1
    End If
    If shp.Cells("User.IndexPers").Result(visNumber) = 36 Then
        wkst.Cells(rowNumber, 4) = shp.Cells("Prop.SetTime").ResultStr(visDate)
        wkst.Cells(rowNumber, 5) = shp.Cells("User.DiameterIn").ResultStr(visUnitsString)
        rowNumber = rowNumber + 1
    End If
    
    
Exit Sub
ex:
    
End Sub
'GFS_Command_13
