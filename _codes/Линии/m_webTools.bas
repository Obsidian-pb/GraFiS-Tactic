Attribute VB_Name = "m_webTools"
Option Explicit
'----------------------��������� �����������-----------------------


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
        Shell "cmd /cstart " & strMain
    Else
        Shell "cmd /cstart http://wiki-fire.org/" & strAlt & ".ashx"
    End If
End Sub


