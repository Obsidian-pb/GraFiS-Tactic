VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Command 
   Caption         =   "�������� �������"
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
Private targetCellName As String









Private Sub UserForm_Activate()
'    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
'
'    Set shp = Application.ActiveWindow.Selection(1)
'
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & delimiter
''    Me.txt_CommandText.Text = GetCurrentTime & delimiter & GetCallName & " "
'    Me.txt_CommandText.Text = GetCurrentTime & delimiter
End Sub

Public Sub NewCommand()
    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    targetCellName = ""
    
    Me.txt_CommandText.text = GetCurrentTime & delimiter & "2" & delimiter
    
    Me.Show
End Sub

Public Sub CurrentCommand(ByRef shp_a As Visio.Shape, ByVal cellName As String)
    Set shp = shp_a
    targetCellName = cellName
    
    Me.txt_CommandText.text = shp.Cells(cellName).ResultStr(visUnitsString)
    
    Me.Show
End Sub

Private Sub btn_Ok_Click()
Dim rowI As Integer
Dim rowT As Integer
Dim rowA As Integer
Dim rowIName As String
Dim tagName As String
Dim frml As String
Dim targetCellNameShort As String
    
    '���� ����� ������� ������ ���, ������ ���������� ��������� ����� ������, ���� ���� - ������������ �� ��� ���������
    If targetCellName = "" Then
'        rowIName = "GFS_Command_" & shp.RowCount(visSectionUser) + 1
        rowIName = "GFS_Command_" & GetNextNumber(shp)
        
        '������ � ������ User
        rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
        shp.CellsSRC(visSectionUser, rowI, 0).Formula = """" & FixText(Me.txt_CommandText) & """"
        
        '������ � ������ ����������
        tagName = "Commands"
        If shp.CellExists("SmartTags.GFS_Commands", False) = 0 Then
            rowT = shp.AddNamedRow(visSectionSmartTag, "GFS_Commands", 0)
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagName).Formula = """" & tagName & """"
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagButtonFace).Formula = 346
        End If
        
        '������ � ������ Action
        rowA = shp.AddNamedRow(visSectionAction, rowIName, 0)
        shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = """" & GetCommandText(75) & """"
        shp.CellsSRC(visSectionAction, rowA, visActionTagName).Formula = """" & tagName & """"
        frml = "CALLTHIS(" & Chr(34) & "RedactThisText" & _
                Chr(34) & "," & Chr(34) & "����������_���" & Chr(34) & "," & _
                Chr(34) & "User." & rowIName & Chr(34) & ")"
        shp.CellsSRC(visSectionAction, rowA, visActionAction).FormulaU = frml
        
        '��������� ������ �������� �������� ������� User.CurrentDocTime, ���� �� ���
        If cellVal(shp, "User.CurrentDocTime", , "-1") < 0 Then
            rowIName = "CurrentDocTime"
            rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
            shp.CellsSRC(visSectionUser, rowI, 0).Formula = "TheDoc!User.CurrentTime"
            frml = "CALLTHIS(" & Chr(34) & "CheckEnd" & _
                Chr(34) & "," & Chr(34) & "����������_���" & Chr(34) & _
                ",User.CurrentDocTime)"
            shp.CellsSRC(visSectionUser, rowI, 1).FormulaU = frml
        End If
        
        '����� � ������� ����� ����� ��������� ������ ����� ������ �������
'        shp.CellsSRC(visSectionAction, rowA, visActionButtonFace).Formula = 346
        
    Else
        shp.Cells(targetCellName).Formula = """" & FixText(Me.txt_CommandText) & """"
        targetCellNameShort = Split(targetCellName, ".")(1)
        shp.Cells("Actions." & targetCellNameShort).Formula = """" & GetCommandText(75) & """"
    End If
    

    
    Me.Hide
End Sub
Private Sub btn_Cancel_Click()
    Me.Hide
End Sub

'������ ��������
Private Sub btn_Delete_Click()
Dim targetCellNameShort As String
Dim targetRowIndex As Integer
    
    ' ���� ������� ������ ���, ������, ����� ������ �������� ��������
    If targetCellName <> "" Then
        targetCellNameShort = Split(targetCellName, ".")(1)
        
        '������ � ������ User
        targetRowIndex = GetRowIndex("User." & targetCellNameShort)
        If targetRowIndex >= 0 Then
            shp.DeleteRow visSectionUser, targetRowIndex
        End If
        '������ � ������ Action
        targetRowIndex = GetRowIndex("Actions." & targetCellNameShort)   'Actions.GFS_Command_14
        If targetRowIndex >= 0 Then
            shp.DeleteRow visSectionAction, targetRowIndex
        End If
    End If
    
    '�������� ������� ����� ��� (��� ������, ���� ������ ������ ���)
    TryDeleteSmartTag "Commands", "SmartTags.GFS_Commands"
    
    Me.Hide
End Sub

Private Function GetCurrentTime() As String
    GetCurrentTime = Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").ResultStr(visUnitsString)
End Function
Private Function FixText(ByVal str As String) As String
    FixText = Replace(str, Chr(34), "'")
End Function
Private Function GetCommandText(Optional ByVal l As Integer = 25) As String
Dim str As String

    On Error GoTo EX

    str = FixText(Me.txt_CommandText)
    If Len(str) < l Then
        GetCommandText = str
    Else
        GetCommandText = left(str, l) & "..."
    End If

Exit Function
EX:
    GetCommandText = "***"
End Function
Private Function GetRowIndex(cellName As String) As Integer
    On Error GoTo EX
    GetRowIndex = shp.Cells(cellName).Row
Exit Function
EX:
    GetRowIndex = -1
End Function

Private Sub TryDeleteSmartTag(stName As String, rowName As String)
'stName - �������� �����-����, rowName - �������� ������ �����-���� � ������ SmartTags
Dim i As Integer
Dim smartTagRowIndex As Integer
    
    '�������� ����� ����� ���� � ��������� ������� �� ����� ����� ���
    smartTagRowIndex = GetRowIndex(rowName)
    If smartTagRowIndex >= 0 Then
        '��������� ��� ������ ������ Actions �� ������� ������� ������ �� ��������� ����� ���
        For i = 0 To shp.RowCount(visSectionAction) - 1
            '���� ���� ���� ���� - ������� �� ��������� �� ������ ����� ���
            If shp.CellsSRC(visSectionAction, i, visActionTagName).ResultStr(visUnitsString) = stName Then Exit Sub
        Next i
    
        shp.DeleteRow visSectionSmartTag, smartTagRowIndex
        '������� ��� �� � ������ ������������ �������
        On Error Resume Next
        shp.DeleteRow visSectionUser, shp.Cells("User.CurrentDocTime").Row
    End If

End Sub
