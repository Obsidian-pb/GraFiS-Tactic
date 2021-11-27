Attribute VB_Name = "Tools"
Option Explicit

Public Function GetScaleAt200() As Double
'���������� ����������� ���������� ������� ������� �������� ������������ �������� 1:200
Dim v_Minor As Double
Dim v_Major As Double

    v_Minor = Application.ActivePage.PageSheet.Cells("PageScale").Result(visNumber)
    v_Major = Application.ActivePage.PageSheet.Cells("DrawingScale").Result(visNumber)
    GetScaleAt200 = (v_Major / v_Minor) / 200
End Function



'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub

