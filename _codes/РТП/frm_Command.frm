VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Command 
   Caption         =   "Добавить команду"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frm_Command.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private shp As Visio.Shape
Const delimiter = " | "






Private Sub UserForm_Activate()
    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    
'    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & delimiter
'    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & " "
    Me.txt_CommandText.Text = GetCurrentTime & delimiter
End Sub


Private Sub btn_Ok_Click()
Dim rowI As Integer
Dim rowT As Integer
Dim rowA As Integer
Dim rowIName As String
Dim tagName As String
    
    rowIName = "GFS_Command_" & shp.RowCount(visSectionUser) + 1
    

    'строка в секции User
    rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
    shp.CellsSRC(visSectionUser, rowI, 0).Formula = """" & FixText(Me.txt_CommandText) & """"
    
    'строка в секции СмартТегов
    tagName = "Commands"
    If shp.CellExists("SmartTags.GFS_Commands", False) = 0 Then
        rowT = shp.AddNamedRow(visSectionSmartTag, "GFS_Commands", 0)
'        shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagDescription).Formula = "User." & rowIName
        shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagName).Formula = """" & tagName & """"
    End If
    
    'строка в секции Action
    rowA = shp.AddNamedRow(visSectionAction, "GFS_Command_" & shp.RowCount(visSectionAction) + 1, 0)
'    shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = """" & GetCommandText(75) & """"
    shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = "User." & rowIName
    shp.CellsSRC(visSectionAction, rowA, visActionTagName).Formula = """" & tagName & """"
    
    Me.Hide
End Sub
Private Sub btn_Cancel_Click()
    Me.Hide
End Sub


Private Function GetCurrentTime() As String
    GetCurrentTime = Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").ResultStr(visUnitsString)
End Function
Private Function GetCallName() As String
    
    On Error GoTo EX
    GetCallName = shp.Cells("Prop.Call").ResultStr(visUnitsString)
Exit Function
EX:
    GetCallName = "-"
End Function
Private Function FixText(ByVal str As String) As String
    FixText = Replace(str, Chr(34), "'")
End Function
Private Function GetCommandText(Optional ByVal l As Integer = 25) As String
Dim str As String
    On Error GoTo EX
    str = Split(Me.txt_CommandText, delimiter)(1)
    If Len(str) < l Then
        GetCommandText = str
    Else
        GetCommandText = Left(str, l) & "..."
    End If
    
Exit Function
EX:
    GetCommandText = "***"
End Function
