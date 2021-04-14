Attribute VB_Name = "m_tools"


Public Function ClearString(ByVal txt As String) As String
'������� �������� �� ���� ������ � ���������� ����������� �� ����� ����� (���� ������ ������������ �����) ��� �������� ������
Dim tmpVal As Variant
    On Error Resume Next
    txt = Round(txt, 2)
    ClearString = txt
End Function

Public Sub sleep(ByVal sec As Single, Optional ByVal doE As Boolean = False)
Dim i As Long
Dim endTime As Single
    
    endTime = DateTime.Timer + sec
    Do While DateTime.Timer < endTime
        If doE Then DoEvents
    Loop
    
'    Debug.Print "sleep 0.5"
    
End Sub

'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.Version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.Description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub

