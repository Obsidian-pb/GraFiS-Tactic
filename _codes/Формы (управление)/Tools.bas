Attribute VB_Name = "Tools"
Option Explicit



Public Function CallToStr(ByVal a_Call As String) As String
'������� �������� ������ � ��������, � ���� �� ����� �������� � ����� - ��������, ���������� ��������� ��������
On Error Resume Next
    a_Call = Int(a_Call)
    CallToStr = a_Call
End Function


Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub

