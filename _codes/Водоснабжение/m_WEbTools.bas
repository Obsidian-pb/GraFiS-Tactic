Attribute VB_Name = "m_WEbTools"
Option Explicit
'----------------------��������� �����������-----------------------
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
'                    ByVal Hwnd As Long, _
'                    ByVal lpOperation As String, _
'                    ByVal lpFile As String, _
'                    ByVal lpParameters As String, _
'                    ByVal lpDirectory As String, _
'                    ByVal nShowCmd As Long) As Long
'
''------------------------------������ � wiki-fire.org-----------------------
'Public Sub GotoWF(ShpObj As Visio.Shape, strMain As String, strAlt As String)
''��������� �� ��������� ����� wiki-fire
''strMain - ������ ����� ���������, ���� ������ - ��������� �� ����
''strAlt - �������������� ����� ��������� - ������������, ���� strMain �� ������
'Const SW_SHOWNORMAL = 1
'
'    '��-��������� ������������� �� �� �������� ������ ������ NULL ���������� �� "0", ������� ������������ ������ �������� �� ������ ������ ""
'    If strMain = "0" Then strMain = ""
'
'    '�������� ����� ����� �� ������ �������������, ��� ������������ ������
'    strAlt = Replace(strAlt, "/", "_")
'
'    If Len(strMain) > 0 Then
'        If InStr(1, strMain, "wiki-fire.org") = 0 Then Exit Sub '���� � ������� ������ ��� �������� �� wiki-fire.org - ���������� ����� �� ��������� - ����� �������� ������ �� ��������� wiki-fire.org
'        ShellExecute 0&, "Open", strMain, _
'                vbNullString, vbNullString, SW_SHOWNORMAL
'    Else
'        ShellExecute 0&, "Open", "http://wiki-fire.org/" & strAlt & ".ashx", _
'                vbNullString, vbNullString, SW_SHOWNORMAL
'    End If
'End Sub


Public Sub GotoWF(ShpObj As Visio.Shape, strMain As String, strAlt As String)
'��������� �� ��������� ����� wiki-fire
'strMain - ������ ����� ���������, ���� ������ - ��������� �� ����
'strAlt - �������������� ����� ��������� - ������������, ���� strMain �� ������
Const SW_SHOWNORMAL = 1
    
    '��-��������� ������������� �� �� �������� ������ ������ NULL ���������� �� "0", ������� ������������ ������ �������� �� ������ ������ ""
    If strMain = "0" Then strMain = ""
    
    '�������� ����� ����� �� ������ �������������, ��� ������������ ������
    strAlt = Replace(strAlt, "/", "_")
    
    '�������� d strMain ������� �� %20
    strMain = Replace(strMain, " ", "%20")
    strAlt = Replace(strAlt, " ", "%20")
    
    If Len(strMain) > 0 Then
        If InStr(1, strMain, "wiki-fire.org") = 0 Then Exit Sub '���� � ������� ������ ��� �������� �� wiki-fire.org - ���������� ����� �� ��������� - ����� �������� ������ �� ��������� wiki-fire.org
'        Shell "cmd /cstart http://www.cbr.ru/"
        Shell "cmd /cstart " & strMain
'        ShellExecute 0&, "Open", strMain, _
                vbNullString, vbNullString, SW_SHOWNORMAL
    Else
        Shell "cmd /cstart http://wiki-fire.org/" & strAlt & ".ashx"
'        ShellExecute 0&, "Open", "http://wiki-fire.org/" & strAlt & ".ashx", _
'                vbNullString, vbNullString, SW_SHOWNORMAL
    End If
End Sub
