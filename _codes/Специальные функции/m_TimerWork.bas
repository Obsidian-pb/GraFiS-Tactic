Attribute VB_Name = "m_TimerWork"
Option Explicit

'Public elpTm As Long
'Public tmrID As Long
'Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

#If VBA7 Then
    Public elpTm As LongPtr
    Public tmrID As LongPtr
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIDEvent As LongPtr, _
                    ByVal uElapse As LongPtr, _
                    ByVal lpTimerfunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIDEvent As LongPtr) As LongPtr
#Else
    Public elpTm As Long
    Public tmrID As Long
    Public Declare Function SetTimer Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal nIDEvent As Long, _
                    ByVal uElapse As Long, _
                    ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal nIDEvent As Long) As Long
#End If

Private timerTB As c_TimerTB
'--------------------------------------------------------------------------------------------------


Public Sub AddTimer()
    On Error Resume Next
    Set timerTB = New c_TimerTB
    TimerButtonChamgeState msoButtonDown
End Sub

Public Sub DelTBTimer()
    On Error Resume Next
    Set timerTB = Nothing
    TimerButtonChamgeState msoButtonUp
End Sub

Public Sub TimerButtonChamgeState(ByVal state As MsoButtonState)
'�������� ��������� ������ "������"
    On Error Resume Next
    Application.CommandBars("�����������").Controls("������").state = state
End Sub

Public Sub AddTBTimer(ShpObj As Visio.Shape)
'---��������� ������� �������
Dim i As Integer

    On Error GoTo ex

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

'---��������� ���� �� ��� ������ ���������� "������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "������" Then 'Exit Sub
            DelTBTimer
            ShpObj.Delete
            Exit Sub
        End If
    Next i

'---������� ������ ���������� "�������� �����"--------------------------------------------
    AddTimer

    ShpObj.Delete

Exit Sub
ex:
    'Error
    ShpObj.Delete

End Sub




'--------------����� �������---------------------------------------------------------------------
Public Sub tmrStart()
'��������� ������ ��� ����� ������� �������
    tmrID = SetTimer(&H0, &H0, 1000, AddressOf tmrPrc)  '1000 - 1 ���
End Sub
'����������� ������� � ��������� ���������,
'�.�. ������ �� ����� "���������" � ��������� ��������� tmrPrc
Public Sub tmrKill()
    KillTimer &H0, tmrID
End Sub

Public Sub tmrPrc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
'��� ���������� ������� - "���������" ���������../
Dim controlDate As CommandBarControl
Dim controlTime As CommandBarControl
Dim timeCell As Visio.Cell
Dim curDateTime As Date

    On Error GoTo ex
'---���������� ������ ���� CurrentDate
    Set controlDate = Application.CommandBars("������").Controls("����")
'---���������� ������ ���� CurrentTime
    Set controlTime = Application.CommandBars("������").Controls("�����")
'---���������� ������ ���������� ������� ����� � ���� �����
    Set timeCell = Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime")
    curDateTime = timeCell.Result(visDate)
    
'    Debug.Print Now() & " - " & controlDate.Text & ", " & controlTime.Text
'---� ������, ���� ���� ������ ������������ "������" ���������� - ��������� ������
    If controlDate.Text = controlTime.Text Then
'        Debug.Print "���������"
        DelTBTimer
        AddTimer
    End If
    
'---���������, �� �������� �� �������� TB_Date � TB_Time
    If Not controlTime.Text = TimeValue(curDateTime) Then
        timerTB.OnCurrentTimeAction
    End If
    If Not controlDate.Text = DateValue(curDateTime) Then
        timerTB.OnCurrentDateAction
    End If
    
'---� ������, ���� ������� ������ �������� ������� - ������������ 10-�� ��������� ������
    If timerTB.CurrentTimerActive Then
        If DateDiff("s", curDateTime, Now()) >= 10 Then
        '---������������� �������� ���� � ������� ��� ���� �����
            controlDate.Text = DateValue(Now())
            controlTime.Text = TimeValue(Now())
        '---�������� ������ � �������
            timerTB.PS_UpdateDateTime controlDate, controlTime
        End If
    End If

'---���� ���� ���������� - ��������� ��
    timerTB.FillFullData

ex:
End Sub


