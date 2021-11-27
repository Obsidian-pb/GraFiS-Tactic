VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_InfoSimple 
   Caption         =   "��������, ����������, ��������� ����������"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frm_InfoSimple.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_InfoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private shp As Visio.Shape
Private targetCellName As String
Private infoType As Byte








Private Sub UserForm_Activate()

End Sub

Public Sub NewInfo(Optional ByVal infoType_a As Byte = 0)
    If Application.ActiveWindow.Selection.Count <> 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    targetCellName = ""
    
    infoType = infoType_a
    
    Select Case infoType
        Case Is = 0     '����������
            Me.txt_InfoText.text = GetCurrentTime & delimiter
        Case Is = 1     '������
            Me.txt_InfoText.text = GetCurrentTime & delimiter & "������" & delimiter
        Case Is = 2     '����������
            Me.txt_InfoText.text = GetCurrentTime & delimiter & "���������" & delimiter
    End Select
    
    
    Me.Show
End Sub

Public Sub CurrentInfo(ByRef shp_a As Visio.Shape, ByVal cellName As String)
    Set shp = shp_a
    targetCellName = cellName
    
    Me.txt_InfoText.text = shp.Cells(cellName).ResultStr(visUnitsString)
    
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
        rowIName = "GFS_Info_" & shp.RowCount(visSectionUser) + 1
        
        '������ � ������ User
        rowI = shp.AddNamedRow(visSectionUser, rowIName, 0)
        shp.CellsSRC(visSectionUser, rowI, 0).Formula = """" & FixText(Me.txt_InfoText) & """"
        
        '������ � ������ ����������
        tagName = "Info"
        If shp.CellExists("SmartTags.GFS_Info", False) = 0 Then
            rowT = shp.AddNamedRow(visSectionSmartTag, "GFS_Info", 0)
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagName).Formula = """" & tagName & """"
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagButtonFace).Formula = 0
            shp.CellsSRC(visSectionSmartTag, rowT, visSmartTagY).Formula = "Height"
        End If
        
        '������ � ������ Action
        rowA = shp.AddNamedRow(visSectionAction, rowIName, 0)
        shp.CellsSRC(visSectionAction, rowA, visActionMenu).Formula = """" & GetCommandText(75) & """"
        shp.CellsSRC(visSectionAction, rowA, visActionTagName).Formula = """" & tagName & """"
        frml = "CALLTHIS(" & Chr(34) & "RedactThisInfo" & _
                Chr(34) & "," & Chr(34) & "���" & Chr(34) & "," & _
                Chr(34) & "User." & rowIName & Chr(34) & ")"
        shp.CellsSRC(visSectionAction, rowA, visActionAction).FormulaU = frml
        
        Select Case infoType
            Case Is = 0     '����������
                shp.CellsSRC(visSectionAction, rowA, visActionButtonFace).FormulaU = 0
            Case Is = 1     '������
                shp.CellsSRC(visSectionAction, rowA, visActionButtonFace).FormulaU = 215
            Case Is = 2     '����������
                shp.CellsSRC(visSectionAction, rowA, visActionButtonFace).FormulaU = 192
        End Select
        
    Else
        shp.Cells(targetCellName).Formula = """" & FixText(Me.txt_InfoText) & """"
        targetCellNameShort = Split(targetCellName, ".")(1)
        shp.Cells("Actions." & targetCellNameShort).Formula = """" & GetCommandText(75) & """"
    End If
    

    
    Me.Hide
End Sub
Private Sub btn_Cancel_Click()
    Me.Hide
End Sub

'������ �������
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
    TryDeleteSmartTag "Info", "SmartTags.GFS_Info"
    
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

    On Error GoTo ex

    str = FixText(Me.txt_InfoText)
    If Len(str) < l Then
        GetCommandText = str
    Else
        GetCommandText = left(str, l) & "..."
    End If

Exit Function
ex:
    GetCommandText = "***"
End Function
Private Function GetRowIndex(cellName As String) As Integer
    On Error GoTo ex
    GetRowIndex = shp.Cells(cellName).Row
Exit Function
ex:
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
    End If

End Sub


