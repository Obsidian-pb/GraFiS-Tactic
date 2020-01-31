Attribute VB_Name = "ManagementFormsUse"
Private c_ManagementTech As c_ManagementTechnics 'Переменная в которой хранитя ссылка на c_ManagementTechnics
Private c_ManagementStvols As c_ManagementStvols 'Переменная в которой хранитя ссылка на c_ManagementStvols
Private c_ManagementGDZS As c_ManagementGDZS 'Переменная в которой хранитя ссылка на c_ManagementGDZS
Private c_ManagementTL As c_ManagementTimeLine 'Переменная в которой хранитя ссылка на c_ManagementTimeLine



'------------------------------------------Процедуры работы с формами менеджмента-------------------------
Public Sub MngmnWndwShow(ShpObj As Visio.Shape)
'Процедура активирует форму ManagementTechnics
    If c_ManagementTech Is Nothing Then
        Set c_ManagementTech = New c_ManagementTechnics
    Else
        c_ManagementTech.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnStvolsWndwShow(ShpObj As Visio.Shape)
'Процедура активирует форму ManagementStvols
    If c_ManagementStvols Is Nothing Then
        Set c_ManagementStvols = New c_ManagementStvols
    Else
        c_ManagementStvols.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnGDZSWndwShow(ShpObj As Visio.Shape)
'Процедура активирует форму ManagementGDZS
    If c_ManagementGDZS Is Nothing Then
        Set c_ManagementGDZS = New c_ManagementGDZS
    Else
        c_ManagementGDZS.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub
Public Sub MngmnTimeLineWndwShow(ShpObj As Visio.Shape)
'Процедура активирует форму ManagementTimeLine
    If c_ManagementTL Is Nothing Then
        Set c_ManagementTL = New c_ManagementTimeLine
    Else
        c_ManagementTL.PS_ShowWindow
    End If
    ShpObj.Delete
End Sub

Public Sub MngmnWndwHide()
'Процедура закрытия формы ManagementTechnics
    Set c_ManagementTech = Nothing
End Sub
Public Sub MngmnStvolsWndwHide()
'Процедура закрытия формы ManagementStvols
    Set c_ManagementStvols = Nothing
End Sub
Public Sub MngmnGDZSWndwHide()
'Процедура закрытия формы ManagementGDZS
    Set c_ManagementGDZS = Nothing
End Sub
Public Sub MngmnTimeLineWndwHide()
'Процедура закрытия формы ManagementGDZS
    Set c_ManagementTL = Nothing
End Sub

Public Sub TryCloseForms()
'Пытаемся закрывать окна (если они были открыты)
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
