VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_TimerTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����� ��� ������ ������ ������������ "������"

Public TB_Date As Office.CommandBarControl  '����
Public TB_Time As Office.CommandBarControl  '�����
Public WithEvents ComButMinuteReduce As Office.CommandBarButton  '��������� �� ������
Attribute ComButMinuteReduce.VB_VarHelpID = -1
Public WithEvents ComButMinuteRize As Office.CommandBarButton  '��������� �� ������
Attribute ComButMinuteRize.VB_VarHelpID = -1
Public WithEvents ComButCurrentTime As Office.CommandBarButton  '������� �����
Attribute ComButCurrentTime.VB_VarHelpID = -1
Public WithEvents ComButSystemTime As Office.CommandBarButton  '��������� �����
Attribute ComButSystemTime.VB_VarHelpID = -1
Public WithEvents ComButFireTime As Office.CommandBarButton  '����� ������ ������
Attribute ComButFireTime.VB_VarHelpID = -1
Public WithEvents ComButLastTime As Office.CommandBarButton  '����� ������� �����
Attribute ComButLastTime.VB_VarHelpID = -1
Public WithEvents ComButShapeTime As Office.CommandBarButton  '����� ��������� ������
Attribute ComButShapeTime.VB_VarHelpID = -1
Public WithEvents ComButTimeToShape As Office.CommandBarButton  '���������� ������� �����
Attribute ComButTimeToShape.VB_VarHelpID = -1

Public CurrentTimerActive As Boolean
'----------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    '---��������� ������ �������
    AddTBTimerDirect
    '---����������� ������
    Set ComButMinuteReduce = Application.CommandBars("������").Controls("��������� �� 1 ������")
    Set ComButMinuteRize = Application.CommandBars("������").Controls("��������� �� 1 ������")
    Set ComButSystemTime = Application.CommandBars("������").Controls("��������� �����")
    Set ComButCurrentTime = Application.CommandBars("������").Controls("������� �����")
    Set ComButFireTime = Application.CommandBars("������").Controls("����� ������ ������")
    Set ComButLastTime = Application.CommandBars("������").Controls("����� ������� �����")
    Set ComButShapeTime = Application.CommandBars("������").Controls("����� ��������� ������")
    Set ComButTimeToShape = Application.CommandBars("������").Controls("���������� ������� �����")
    '---����������� ����
    Set TB_Date = Application.CommandBars("������").Controls("����")
    Set TB_Time = Application.CommandBars("������").Controls("�����")
    '---��������� ������ ��-���������
    CurrentTimerActive = False
    '---�������� ����� ������
    tmrStart
End Sub

Private Sub Class_Terminate()
    '---������������ ������
    RemoveTBTimer
End Sub

'----------------------------------------------------------------------------------------
Public Sub AddTBTimerDirect()
'---��������� ������� �������
Dim i As Integer

    On Error GoTo ex

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

'---��������� ���� �� ��� ������ ���������� "������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "������" Then Exit Sub
    Next i

'---������� ������ ���������� "�������� �����"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True)
    With Bar
        .Name = "������"
        .Visible = True
    End With

    AddTBTimerButtons

Exit Sub
ex:
    'Error

End Sub

Public Sub RemoveTBTimer()
'��������� ���������� ������ ���������� "������"-------------------------------
    On Error GoTo ex
'---��������� ������ �������� ���������� �������, ���� �� ��� �������
    tmrKill
'---������� ������ ����������
    Application.CommandBars("������").Delete
ex:
End Sub

Public Sub AddTBTimerButtons()
'---��������� ������ �� ������� �������
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, control As CommandBarControl, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("������")
    
    Dim tbc As c_TimerTB
    
'---��������� ������ �� ������ ���������� "������"--------------------------------
'---������� "����"
    Set control = Bar.Controls.Add(Type:=msoControlEdit)
    With control
        .Caption = "����"
        .Tag = "CurrentDate"
        .TooltipText = "������� ���� �����"
        .Width = 55
'        .OnAction = "OnCurrentDateAction"
    End With
    PS_LoadFromSheetDate control
    Set control = Nothing
'---������ "��������� �� ������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������� �� 1 ������"
        .Tag = "Decrease1Min"
        .TooltipText = "��������� ������� ����� ����� �� 1 ������"
        .FaceID = 2944
        .BeginGroup = True
    End With
    Set Button = Nothing
'---������� "����"
    Set control = Bar.Controls.Add(Type:=msoControlEdit)
    With control
        .Caption = "�����"
        .Tag = "CurrentTime"
        .TooltipText = "������� ����� �����"
        .Width = 45
    End With
    PS_LoadFromSheetTime control
    Set control = Nothing
'---������ "��������� �� ������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������� �� 1 ������"
        .Tag = "Rise1min"
        .TooltipText = "��������� ������� ����� ����� �� 1 ������"
        .FaceID = 2945
    End With
    Set Button = Nothing
'---������ "��������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������� �����"
        .Tag = "SystemTime"
        .TooltipText = "������������� �� ���������� ����� (������)"
        .FaceID = 2743
        .BeginGroup = True
'        .OnAction = "OnSystemTimeAction"
    End With
    Set Button = Nothing
'---������ "������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������� �����"
        .Tag = "CurrentTimeContinous"
        .TooltipText = "������������� �� ������� ���������� ����� (������ � ���������� 10�.)"
        .FaceID = 2146
    End With
    Set Button = Nothing
'---������ "����� ������ ������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����� ������ ������"
        .Tag = "FireTime"
        .TooltipText = "������������� �� ����� ������ ������"
        .FaceID = 1885
    End With
    Set Button = Nothing
'---������ "����� ������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����� ������� �����"
        .Tag = "LastTime"
        .TooltipText = "������������� �� ����� ����� ������� ������"
        .FaceID = 157
    End With
    Set Button = Nothing
'---������ "����� ��������� ������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����� ��������� ������"
        .Tag = "ShapeTime"
        .TooltipText = "������������� �� ����� ��������� ������"
        .FaceID = 110
    End With
    Set Button = Nothing
'---������ "���������� ������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "���������� ������� �����"
        .Tag = "TimeToShape"
        .TooltipText = "���������� ��� ��������� ������ ������� ����� �����"
        .FaceID = 111
        .BeginGroup = True
    End With
    Set Button = Nothing


Set Bar = Nothing
Exit Sub
ex:
'    Debug.Print Err.Description
    Set Bar = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddTBTimerButtons"
End Sub

Public Sub RemooveTBTimer()
'��������� ���������� ������ ���������� "������"
    Application.CommandBars("������").Delete
End Sub




'-------------------------����������� ������� ������----------------------------------------------------
Private Sub ComButMinuteReduce_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnDelMinuteAction
End Sub
Private Sub ComButMinuteRize_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnAddMinuteAction
End Sub
Private Sub ComButSystemTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnSystemTimeAction
End Sub
Private Sub ComButCurrentTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnCurrentTimeComtinousAction
End Sub
Private Sub ComButFireTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnFireTimeAction
End Sub
Private Sub ComButLastTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnLastTimeAction
End Sub
Private Sub ComButShapeTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnShapeTimeAction
End Sub
Private Sub ComButTimeToShape_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnShapeTimeToShapeAction
End Sub






'-------------------------����� � ������� �� ������������ �� �������� � ����������----------------------
Public Sub OnCurrentDateAction()
'����� ����������� ��� ������� ������ � ���� CurrentDate
'---��������� ������ (���� ��� �������)
'    tmrKill
    On Error GoTo ex
'---��������� ������������ ���������� ��������
    If IsDate(TB_Date.Text) Then
    '---���� ��������� - ��������� �������� ������ TheDoc!User.CurrentTime
        PS_UpdateDateTime TB_Date, TB_Time
    Else
    '---���� ����������� - ���������� ��, ��� ��������� � ������ TheDoc!User.CurrentTime
        PS_LoadFromSheetDate TB_Date
    End If
    
ex:
End Sub

Public Sub OnCurrentTimeAction()
'����� ����������� ��� ������� ������ � ���� CurrentTime
    
'---��������� ������ (���� ��� �������)
'    tmrKill
    On Error GoTo ex
'---��������� ������������ ���������� ��������
    If IsDate(TB_Time.Text) Then
    '---���� ��������� - ��������� �������� ������ TheDoc!User.CurrentTime
        PS_UpdateDateTime TB_Date, TB_Time
    Else
    '---���� ����������� - ���������� ��, ��� ��������� � ������ TheDoc!User.CurrentTime
        PS_LoadFromSheetTime TB_Time
    End If
    
ex:
End Sub

Private Sub OnDelMinuteAction()
'����� ����������� ��� ������� �� ������ "��������� �� 1 ������"
Dim tmpDate As Date
    
    On Error GoTo ex
    
'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False

'---�������� ������� ��������
    tmpDate = TimeValue(TB_Time.Text)
    
'---��������� ��� �� 1 ������
    tmpDate = DateAdd("n", -1, tmpDate)
    
'---������������� ����� �������� ��� ��������
    TB_Time.Text = TimeValue(tmpDate)
    
'---���� ����� �������� = 59����� �� ��������� ���� �� 1 ����
    If Hour(tmpDate) = 23 And Minute(tmpDate) = 59 Then
        PS_DateChange TB_Date, -1
    End If
    
'---�������� ������ � �������
    PS_UpdateDateTime TB_Date, TB_Time

ex:
End Sub
Private Sub OnAddMinuteAction()
'����� ����������� ��� ������� �� ������ "��������� �� 1 ������"
Dim tmpDate As Date

    On Error GoTo ex

'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False

'---�������� ������� ��������
    tmpDate = TimeValue(TB_Time.Text)
    
'---��������� ��� �� 1 ������
    tmpDate = DateAdd("n", 1, tmpDate)
    
'---������������� ����� �������� ��� ��������
    TB_Time.Text = TimeValue(tmpDate)
    
'---���� ����� �������� = 59����� �� ��������� ���� �� 1 ����
    If Hour(tmpDate) = 0 And Minute(tmpDate) = 0 Then
        PS_DateChange TB_Date, 1
    End If
    
'---�������� ������ � �������
    PS_UpdateDateTime TB_Date, TB_Time

ex:
End Sub

Private Sub OnSystemTimeAction()
'����� �������������� ������� �� ������ "������������� �� ��������� �����"

    On Error GoTo ex

'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False

'---������������� �������� ���� � ������� ��� ���� �����
    TB_Date.Text = DateValue(Now())
    TB_Time.Text = TimeValue(Now())
    
'---�������� ������ � �������
    PS_UpdateDateTime TB_Date, TB_Time
    
ex:
End Sub

Private Sub OnFireTimeAction()
'����� ����������� ��� ������� �� ������ "FireTime"

    On Error GoTo ex

'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False
    
'---������������� ������ ��� ������ "TheDoc!User.CurrentTime"
    If GetFireTime Then
        Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
            "User.FireTime"
    Else
        MsgBox "�� ������� ������� ����� ��� ����� ����� �������"
        Exit Sub
    End If

    
'---������������� �������� ���� � ������� ��� ���� �����
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnCurrentTimeComtinousAction()
'����� ����������� ��� ������� �� ������ "CurrentTime"

    On Error GoTo ex

'---���������� ������ �������� �������
    OnSystemTimeAction

'---�������� ������ (���� ��� ��������)
'    tmrStart
    CurrentTimerActive = True
    
ex:
End Sub

Private Sub OnLastTimeAction()
'����� ����������� ��� ������� �� ������ "FireTime"

    On Error GoTo ex

'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False
    
'---������������� ������ ��� ������ "TheDoc!User.CurrentTime"
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        GetLastTime(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.FireTime").Result(visDate))
    
'---������������� �������� ���� � ������� ��� ���� �����
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnShapeTimeAction()
'����� ����������� ��� ������� �� ������ "ShapeTime" - �������� ����� ��������� ������

    On Error GoTo ex
    
'---���� �� ���� ������ �� �������� - �������
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
'---��������� ������ (���� ��� �������)
    CurrentTimerActive = False
    
'---������������� �������� ��� ������ "TheDoc!User.CurrentTime"
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        GetCurrentShapeTime(Application.ActiveWindow.Selection(1))
    
'---������������� �������� ���� � ������� ��� ���� �����
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnShapeTimeToShapeAction()
'����� ����������� ��� ������� �� ������ "TimeToShape" - ���������� ��� ��������� ������ (�����) ������� ����� �����
Dim shp As Visio.Shape
Dim curTime As Double

    On Error GoTo ex
    
'---���� �� ���� ������ �� �������� - �������
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
'---�������� ������� ����� �����
    curTime = Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Result(visDate)
    
'---���������� ��� ������ � ���������
    For Each shp In Application.ActiveWindow.Selection
        SetCurrentShapeTime shp, curTime
    Next shp
    
ex:
End Sub


'---------------��������� �����------------------------------------------------------------------
Private Sub PS_DateChange(ByRef controlDate As CommandBarControl, ByVal changeValue As Double)
'����� ��������� �������� ���� � �������� CurrentDate �� 1 ����
Dim tmpDate As Date

    On Error GoTo ex

'---�������� ������� ��������
    tmpDate = DateValue(controlDate.Text)
    
'---��������� ��� �� 1 ����
    tmpDate = DateAdd("d", changeValue, tmpDate)
    
'---������������� ����� �������� ��� ��������
    controlDate.Text = DateValue(tmpDate)
    
ex:
End Sub

Public Sub FillFullData()
'���� � ����� ������� � ���� ����� - �������������

    On Error GoTo ex

'    Debug.Print TB_Date.Text
    If TB_Date.Text = "" Then
        PS_LoadFromSheetDate TB_Date
    End If
    If TB_Time.Text = "" Then
        PS_LoadFromSheetTime TB_Time
    End If
    
ex:
End Sub

Private Function GetLastTime(ByVal fireStartTime As Double) As String
'������� ���������� �������� ���������� �������������� �������
Dim shp As Visio.Shape
Dim curVal As Double
Dim maxVal As Double
    
    On Error GoTo ex
    
    maxVal = fireStartTime
    '---���������� ��� ������ �� ����� � ���� �������� ������ �� ������� ������� ������������ ���������� ���
    For Each shp In Application.ActivePage.Shapes
        curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
                 GetShapeTime(shp, "Prop.LineTime") + _
                 GetShapeTime(shp, "Prop.SetTime") + _
                 GetShapeTime(shp, "Prop.SquareTime") + _
                 GetShapeTime(shp, "Prop.StabCreationTime") + _
                 GetShapeTime(shp, "Prop.UTPCreationTime") + _
                 GetShapeTime(shp, "Prop.FormingTime") + _
                 GetShapeTime(shp, "Prop.FindTime") + _
                 GetShapeTime(shp, "Prop.RushTime")
        
        If curVal > maxVal And curVal < maxVal + 100 Then maxVal = curVal '����� +100 -> ���� �������� ��� ����� � ������� ��������� ����� � ������� ������, �������� ������ ����
        
        '---������������� ������ ��� ������� (���������) ������� ������ ������. ��������� �� ������ ����, �� ��������� ���������� ����� ����� �������������, �� 1, �.�. � �������� ������� ��� 0-� ����
        If maxVal <= 1 Then maxVal = curVal
    Next shp
    
    
    
    GetLastTime = "DATETIME(" & str(maxVal) & ")"
Exit Function
ex:
    GetLastTime = "User.FireTime"
End Function
Private Function GetShapeTime(ByRef shp As Visio.Shape, ByVal cellName) As Double
'������� ���������� �������� ��� ������� ������, ���� ����� ������ ��� - ������������ 0
On Error GoTo ex
    GetShapeTime = shp.Cells(cellName).Result(visDate)
    If GetShapeTime = 0 Then
        GetShapeTime = DateValue(shp.Cells(cellName).ResultStr(visUnitsString)) + _
                       TimeValue(shp.Cells(cellName).ResultStr(visUnitsString))
    End If
Exit Function
ex:
    GetShapeTime = 0
End Function

Private Function GetCurrentShapeTime(ByRef shp As Visio.Shape) As String
'������� ���������� �������� ������� ��� ��������� ������
Dim curVal As Double
    
    On Error GoTo ex

'---�������� �������� ����� ������ ���������� �������� ������� �������� (��������� � �.�.)
    curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
         GetShapeTime(shp, "Prop.LineTime") + _
         GetShapeTime(shp, "Prop.SetTime") + _
         GetShapeTime(shp, "Prop.SquareTime") + _
         GetShapeTime(shp, "Prop.StabCreationTime") + _
         GetShapeTime(shp, "Prop.UTPCreationTime") + _
         GetShapeTime(shp, "Prop.FormingTime") + _
         GetShapeTime(shp, "Prop.FindTime") + _
         GetShapeTime(shp, "Prop.RushTime") + _
         GetShapeTime(shp, "Prop.FireTime")
    
    curVal = GetGDZSChainCorrectTime(shp, curVal)
    
    If curVal = 0 Then
        MsgBox "������ ������ �� �������� ���������� �������! �������� �������� ������� ����� ���������� ������ ������� ������ ������.", vbInformation, "��������������"
        GetCurrentShapeTime = "User.FireTime"
    Else
        GetCurrentShapeTime = "DATETIME(" & str(curVal) & ")"
    End If

Exit Function
ex:
    GetCurrentShapeTime = "User.FireTime"
End Function

Private Sub SetCurrentShapeTime(ByRef shp As Visio.Shape, ByRef dateVal As Double)
'��������� ������������� �������� ������� ��� ��������� ������
    On Error Resume Next
    With shp
        .Cells("Prop.ArrivalTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.LineTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.SetTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.SquareTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.StabCreationTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.UTPCreationTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FormingTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FindTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.RushTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FireTime").Formula = "DATETIME(" & dateVal & ")"
    End With
 
End Sub

Private Function GetGDZSChainCorrectTime(ByRef shp As Visio.Shape, ByVal testVal As Double) As Double
'������� ��������� ����������� �������� ��� ����� ������� ����
On Error GoTo ex
    If shp.CellExists("User.IndexPers", 0) Then
        If shp.Cells("User.IndexPers").Result(visNumber) = 46 Then
            GetGDZSChainCorrectTime = shp.Cells("Prop.FormingTime").Result(visDate)
            Exit Function
        End If
    End If
ex:
    GetGDZSChainCorrectTime = testVal
End Function

Private Function GetFireTime() As Boolean
'������� ������������� �������� ������� ������ ������ ��� ������ User.FireTime � ���������� ������ ���� ������ ����
'���� - ���� ���
Dim shp As Visio.Shape
Dim timeVal As Double
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers").Result(visNumber) = 70 Then
                timeVal = GetShapeTime(shp, "Prop.FireTime")
                Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.FireTime").Formula = _
                    "DATETIME(" & str(timeVal) & ")"
                
                GetFireTime = True
                Exit Function
            End If
        End If
    Next shp

GetFireTime = False
Exit Function
ex:
    GetFireTime = False
End Function



'---------------����� ������ �������------------------------------------------------------------
Public Sub PS_UpdateDateTime(ByRef controlDate As CommandBarControl, ByRef controlTime As CommandBarControl)
'��������� �������� ���� TheDoc!User.CurrentTime � ����������� � ���������� ���������
Dim tmpDate As Date
    On Error GoTo ex
    tmpDate = DateValue(controlDate.Text) + TimeValue(controlTime.Text)
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        "DATETIME(" & str(CDbl(tmpDate)) & ")"

ex:
End Sub

Public Sub PS_LoadFromSheetDate(ByRef control As CommandBarControl)
'������������� ���� �������� ���� �������� ������� ���� � ������ TheDoc!User.CurrentTime (�������� ���������)
    On Error Resume Next
    control.Text = _
            Left(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").ResultStr(0), 10)
End Sub

Public Sub PS_LoadFromSheetTime(ByRef control As CommandBarControl)
'������������� ���� �������� ������� �������� ������� ���� � ������ TheDoc!User.CurrentTime (�������� ���������)
    On Error Resume Next
    control.Text = _
            Mid(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").ResultStr(0), 12)
End Sub





'Private Sub VisioApp_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
''Debug.Print GetKeyState(16)
''GetKeyState (VKey)
'End Sub
'
''-------------���� ������� ��������� ������----------------------
'Private Sub VisioApp_SelectionChanged(ByVal Window As IVWindow)
''��������� �� ��������� Selection
'Dim val As Integer
'On Error Resume Next
'    val = GetKeyState(16)
''Debug.Print val
''    If Window.Selection.Count = 0 Then Exit Sub
'    If val <> 0 Then OnSelectionChangeAction Window.Selection(1)
'End Sub
'
'Private Sub OnSelectionChangeAction(ByRef shp As Visio.Shape)
''����� ����������� ��� ������� �� ������ "FireTime"
'Dim curVal As Double
'
'    On Error GoTo EX
'
''---��������� ������ (���� ��� �������)
'    CurrentTimerActive = False
'
''---�������� ����� ������ ������
'    curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
'         GetShapeTime(shp, "Prop.LineTime") + _
'         GetShapeTime(shp, "Prop.SetTime") + _
'         GetShapeTime(shp, "Prop.SquareTime") + _
'         GetShapeTime(shp, "Prop.StabCreationTime") + _
'         GetShapeTime(shp, "Prop.UTPCreationTime") + _
'         GetShapeTime(shp, "Prop.FormingTime") + _
'         GetShapeTime(shp, "Prop.FindTime") + _
'         GetShapeTime(shp, "Prop.RushTime")
'
''---������������� ������ ��� ������ "TheDoc!User.CurrentTime"
'    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
'        "DATETIME(" & Str(curVal) & ")"
'
''---������������� �������� ���� � ������� ��� ���� �����
'    PS_LoadFromSheetDate TB_Date
'    PS_LoadFromSheetTime TB_Time
'
'Exit Sub
'EX:
'    Debug.Print Err.Description
'End Sub

