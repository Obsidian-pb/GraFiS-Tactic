Attribute VB_Name = "ManagementFormsUse"
Private c_ManagementTech As c_ManagementTechnics '���������� � ������� ������� ������ �� c_ManagementTechnics
Private c_ManagementStvols As c_ManagementStvols '���������� � ������� ������� ������ �� c_ManagementStvols
Private c_ManagementGDZS As c_ManagementGDZS '���������� � ������� ������� ������ �� c_ManagementGDZS
Private c_ManagementTL As c_ManagementTimeLine '���������� � ������� ������� ������ �� c_ManagementTimeLine



'------------------------------------------��������� ������ � ������� �����������-------------------------
Public Sub MngmnWndwShow(ShpObj As Visio.Shape)
'��������� ���������� ����� ManagementTechnics
    If c_ManagementTech Is Nothing Then
        Set c_ManagementTech = New c_ManagementTechnics
    Else
        c_ManagementTech.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnStvolsWndwShow(ShpObj As Visio.Shape)
'��������� ���������� ����� ManagementStvols
    If c_ManagementStvols Is Nothing Then
        Set c_ManagementStvols = New c_ManagementStvols
    Else
        c_ManagementStvols.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnGDZSWndwShow(ShpObj As Visio.Shape)
'��������� ���������� ����� ManagementGDZS
    If c_ManagementGDZS Is Nothing Then
        Set c_ManagementGDZS = New c_ManagementGDZS
    Else
        c_ManagementGDZS.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnTimeLineWndwShow(ShpObj As Visio.Shape)
'��������� ���������� ����� ManagementTimeLine
    If c_ManagementTL Is Nothing Then
        Set c_ManagementTL = New c_ManagementTimeLine
    Else
        c_ManagementTL.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub

Public Sub MngmnWndwHide()
'��������� �������� ����� ManagementTechnics
    Set c_ManagementTech = Nothing
End Sub
Public Sub MngmnStvolsWndwHide()
'��������� �������� ����� ManagementStvols
    Set c_ManagementStvols = Nothing
End Sub
Public Sub MngmnGDZSWndwHide()
'��������� �������� ����� ManagementGDZS
    Set c_ManagementGDZS = Nothing
End Sub
Public Sub MngmnTimeLineWndwHide()
'��������� �������� ����� ManagementGDZS
    Set c_ManagementTL = Nothing
End Sub

Public Sub TryCloseForms()
'�������� ��������� ���� (���� ��� ���� �������)
On Error Resume Next
    If Not c_ManagementTech Is Nothing Then
        c_ManagementTech.CloseWindow
        Set c_ManagementTech = Nothing
    End If
    If Not c_ManagementStvols Is Nothing Then
        c_ManagementStvols.CloseWindow
        Set c_ManagementStvols = Nothing
    End If
    If Not c_ManagementGDZS Is Nothing Then
        c_ManagementGDZS.CloseWindow
        Set c_ManagementGDZS = Nothing
    End If
    If Not c_ManagementTL Is Nothing Then
        c_ManagementTL.CloseWindow
        Set c_ManagementTL = Nothing
    End If
End Sub
