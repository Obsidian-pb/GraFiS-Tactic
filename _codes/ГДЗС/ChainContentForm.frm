VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChainContentForm 
   Caption         =   "����������� ����������� ������ � �����"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   OleObjectBlob   =   "ChainContentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChainContentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VS_DevceType As String
Public VS_DeviceModel As String
Public VL_InitShapeID As Long
Public VB_TimeChange As Boolean
Public VB_TimeArrivalChange As Boolean






'------------------------------------------������� ��������� �����----------------------------------------------
Private Sub CB_Conditions_Change()
    ps_ResultsRecalc
End Sub



Private Sub TB_DirectExpense_Change()
    On Error GoTo EX
    TB_DirectExpense.Value = CInt(TB_DirectExpense.Value)
'    If Not (TB_DirectExpense = "" Or TB_DirectExpense = 0) Then ps_ResultsRecalc
    ps_ResultsRecalc
EX:
    Exit Sub
End Sub

Private Sub TB_MainTimeEnter_Change()
Dim vDt_Time As Date
    On Error GoTo EX
    Me.VB_TimeChange = True
    vDt_Time = _
        DateAdd("s", CDbl(TB_TimeTotal.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
    Me.TB_MainTimeExit.Value = _
        IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
        & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
        & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
EX:
    Exit Sub
End Sub

Private Sub TB_TimeArrival_Change()
'��������� ������� ��������� ������� �������� � �����
Dim vDt_Time As Date
    On Error GoTo EX
    Me.VB_TimeArrivalChange = True
    '---��������� ����� ������ ������� � ������, ���� ���� ���������
        vDt_Time = _
            DateAdd("s", CDbl(Me.TB_TimeAtFire.Value) * 60, TimeValue(Me.TB_TimeArrival))
        Me.TB_OrderTimeAtFire.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
EX:
    Exit Sub

End Sub

Private Sub UserForm_Activate()
    
    ps_Conditions_ListFill
    ps_Models_ListFill

    ps_ResultsRecalc
    
End Sub

Private Sub CB_Quit_Click()
    Me.Hide
End Sub

Private Sub CB_OK_Click()
    ps_ValuesBack
    Me.Hide
End Sub

'------------------------------------------��������� ��������� �����----------------------------------------------
Private Sub ps_Conditions_ListFill()
'��������� ���������� ����������� ������ ������� ������
    With CB_Conditions
        .Clear
        .AddItem "����������� �������"
        .AddItem "������� �������"
        .ListIndex = 0
    End With

End Sub

Private Sub ps_Models_ListFill()
'��������� ��������� ������ �� ��������� � ���������� �� �� ������ ������
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim vS_Path As String
Dim i As Integer
    
    On Error GoTo EX
    
'---������� ������ �� ������� ��������
LB_DeviceModel.Clear

'!!!����� �������!!!
'VS_DevceType = "����"

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������ � ����������� �� ���� ���������
        If VS_DevceType = "����" Then
'            SQLQuery = "SELECT ����.������, ����.[����� ��������], ����.[�������� ���������], �������.��� " _
'            & "FROM ������� LEFT JOIN ���� ON �������.���������� = ����.������ ORDER BY ����.������;"
            SQLQuery = "SELECT ����.������, ����.[����� ��������], ����.[�������� ���������], �������.��� " _
            & "FROM ������� RIGHT JOIN ���� ON �������.���������� = ����.������ ORDER BY ����.������;"
        Else
            SQLQuery = "SELECT ����.������, ����.[����� ��������], ����.[�������� ���������], ����.��� " _
            & "FROM ����;"
        End If
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
        
    '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
    With rst
        .MoveFirst
        i = 0
        Do Until .EOF
            LB_DeviceModel.AddItem
            LB_DeviceModel.Column(0, i) = !������
            LB_DeviceModel.Column(1, i) = CStr(![����� ��������])
            LB_DeviceModel.Column(2, i) = ![�������� ���������]
            LB_DeviceModel.Column(3, i) = CStr(!���)
            i = i + 1
            .MoveNext
        Loop
    End With
    
'    LB_DeviceModel.ListIndex = 0
    For i = 0 To Me.LB_DeviceModel.ListCount - 1
'        Me.LB_DeviceModel.ListIndex = i
        If Me.LB_DeviceModel.Column(0, i) = VS_DeviceModel Then
            Me.LB_DeviceModel.Value = i
            Exit For
        End If
    Next i
    
'---������� ��������� ����������
    pth = ""
    Set dbs = Nothing
    Set rst = Nothing
   
Exit Sub
EX:
    SaveLog Err, "ps_Models_ListFill"
End Sub

Private Sub LB_DeviceModel_Change()
    If LB_DeviceModel.ListCount > 0 Then
        TB_BallonsValue = LB_DeviceModel.Column(1, LB_DeviceModel.Value)
        TB_ReductorNeedPressure = LB_DeviceModel.Column(2, LB_DeviceModel.Value)
        TB_CompFactor = LB_DeviceModel.Column(3, LB_DeviceModel.Value)
        '---��������� ���� ����������� �������
        ps_ResultsRecalc
    End If
End Sub



Private Sub ChkB_Perc3_Change()
    TB_Perc3.Enabled = ChkB_Perc3.Value
    TB_Perc3_P1.Enabled = ChkB_Perc3.Value
    TB_Perc3_P2.Enabled = ChkB_Perc3.Value
    
    If ChkB_Perc3.Value = True Then
        TB_Perc3.BackStyle = fmBackStyleOpaque
        TB_Perc3_P1.BackStyle = fmBackStyleOpaque
        TB_Perc3_P2.BackStyle = fmBackStyleOpaque
        TB_Perc3_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc3.BackStyle = fmBackStyleTransparent
        TB_Perc3_P1.BackStyle = fmBackStyleTransparent
        TB_Perc3_P2.BackStyle = fmBackStyleTransparent
        TB_Perc3_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc4_Change()
    TB_Perc4.Enabled = ChkB_Perc4.Value
    TB_Perc4_P1.Enabled = ChkB_Perc4.Value
    TB_Perc4_P2.Enabled = ChkB_Perc4.Value

    If ChkB_Perc4.Value = True Then
        TB_Perc4.BackStyle = fmBackStyleOpaque
        TB_Perc4_P1.BackStyle = fmBackStyleOpaque
        TB_Perc4_P2.BackStyle = fmBackStyleOpaque
        TB_Perc4_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc4.BackStyle = fmBackStyleTransparent
        TB_Perc4_P1.BackStyle = fmBackStyleTransparent
        TB_Perc4_P2.BackStyle = fmBackStyleTransparent
        TB_Perc4_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc5_Change()
    TB_Perc5.Enabled = ChkB_Perc5.Value
    TB_Perc5_P1.Enabled = ChkB_Perc5.Value
    TB_Perc5_P2.Enabled = ChkB_Perc5.Value

    If ChkB_Perc5.Value = True Then
        TB_Perc5.BackStyle = fmBackStyleOpaque
        TB_Perc5_P1.BackStyle = fmBackStyleOpaque
        TB_Perc5_P2.BackStyle = fmBackStyleOpaque
        TB_Perc5_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc5.BackStyle = fmBackStyleTransparent
        TB_Perc5_P1.BackStyle = fmBackStyleTransparent
        TB_Perc5_P2.BackStyle = fmBackStyleTransparent
        TB_Perc5_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

Private Sub ChkB_Perc6_Change()
    TB_Perc6.Enabled = ChkB_Perc6.Value
    TB_Perc6_P1.Enabled = ChkB_Perc6.Value
    TB_Perc6_P2.Enabled = ChkB_Perc6.Value

    If ChkB_Perc6.Value = True Then
        TB_Perc6.BackStyle = fmBackStyleOpaque
        TB_Perc6_P1.BackStyle = fmBackStyleOpaque
        TB_Perc6_P2.BackStyle = fmBackStyleOpaque
        TB_Perc6_PFall.BackStyle = fmBackStyleOpaque
    Else
        TB_Perc6.BackStyle = fmBackStyleTransparent
        TB_Perc6_P1.BackStyle = fmBackStyleTransparent
        TB_Perc6_P2.BackStyle = fmBackStyleTransparent
        TB_Perc6_PFall.BackStyle = fmBackStyleTransparent
    End If
    ps_ResultsRecalc
End Sub

'------------------------------------------��������� ���������� ����������� ��������--------------------------------
Private Sub TB_Perc1_P1_Change()
    If Not TB_Perc1_P1 = "" Then TB_Perc1_PFall = TB_Perc1_P1 - TB_Perc1_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc1_P2_Change()
    If Not TB_Perc1_P2 = "" Then TB_Perc1_PFall = TB_Perc1_P1 - TB_Perc1_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc2_P1_Change()
    If Not TB_Perc2_P1 = "" Then TB_Perc2_PFall = TB_Perc2_P1 - TB_Perc2_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc2_P2_Change()
    If Not TB_Perc2_P2 = "" Then TB_Perc2_PFall = TB_Perc2_P1 - TB_Perc2_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc3_P1_Change()
    If Not TB_Perc3_P1 = "" Then TB_Perc3_PFall = TB_Perc3_P1 - TB_Perc3_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc3_P2_Change()
    If Not TB_Perc3_P2 = "" Then TB_Perc3_PFall = TB_Perc3_P1 - TB_Perc3_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc4_P1_Change()
    If Not TB_Perc4_P1 = "" Then TB_Perc4_PFall = TB_Perc4_P1 - TB_Perc4_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc4_P2_Change()
    If Not TB_Perc4_P2 = "" Then TB_Perc4_PFall = TB_Perc4_P1 - TB_Perc4_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc5_P1_Change()
    If Not TB_Perc5_P1 = "" Then TB_Perc5_PFall = TB_Perc5_P1 - TB_Perc5_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc5_P2_Change()
    If Not TB_Perc5_P2 = "" Then TB_Perc5_PFall = TB_Perc5_P1 - TB_Perc5_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc6_P1_Change()
    If Not TB_Perc6_P1 = "" Then TB_Perc6_PFall = TB_Perc6_P1 - TB_Perc6_P2
    ps_ResultsRecalc
End Sub

Private Sub TB_Perc6_P2_Change()
    If Not TB_Perc6_P2 = "" Then TB_Perc6_PFall = TB_Perc6_P1 - TB_Perc6_P2
    ps_ResultsRecalc
End Sub

Private Sub Min_P1()
'��������� ���������� ������������ �������� ��� ���������
Dim x(6) As Integer
Dim i As Integer
Dim min As Integer

    If Not (TB_Perc1_P1 = "" Or TB_Perc1_P1 = 0) Then x(0) = TB_Perc1_P1 Else x(0) = 300
    If Not (TB_Perc2_P1 = "" Or TB_Perc2_P1 = 0) Then x(1) = TB_Perc2_P1 Else x(1) = 300
    If Not (TB_Perc3_P1 = "" Or TB_Perc3_P1 = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_P1 Else x(2) = 300
    If Not (TB_Perc4_P1 = "" Or TB_Perc4_P1 = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_P1 Else x(3) = 300
    If Not (TB_Perc5_P1 = "" Or TB_Perc5_P1 = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_P1 Else x(4) = 300
    If Not (TB_Perc6_P1 = "" Or TB_Perc6_P1 = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_P1 Else x(5) = 300
    
    min = x(0)
    For i = 0 To 5
        If x(i) < min Then min = x(i)
    Next i
    
    TB_P1_Min = min
End Sub

Private Sub Min_P2()
'��������� ���������� ������������ �������� � ����� ������
Dim x(6) As Integer
Dim i As Integer
Dim min As Integer

    If Not (TB_Perc1_P2 = "" Or TB_Perc1_P2 = 0) Then x(0) = TB_Perc1_P2 Else x(0) = 300
    If Not (TB_Perc2_P2 = "" Or TB_Perc2_P2 = 0) Then x(1) = TB_Perc2_P2 Else x(1) = 300
    If Not (TB_Perc3_P2 = "" Or TB_Perc3_P2 = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_P2 Else x(2) = 300
    If Not (TB_Perc4_P2 = "" Or TB_Perc4_P2 = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_P2 Else x(3) = 300
    If Not (TB_Perc5_P2 = "" Or TB_Perc5_P2 = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_P2 Else x(4) = 300
    If Not (TB_Perc6_P2 = "" Or TB_Perc6_P2 = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_P2 Else x(5) = 300
    
    min = x(0)
    For i = 0 To 5
        If x(i) < min Then min = x(i)
    Next i
    
    TB_P2_Min = min
End Sub

Private Sub Max_PFall()
'��������� ���������� ������������� ������� �������� �� ���� � ����� ������
Dim x(6) As Integer
Dim i As Integer
Dim max As Integer

    If Not (TB_Perc1_PFall = "" Or TB_Perc1_PFall = 0) Then x(0) = TB_Perc1_PFall Else x(0) = 0
    If Not (TB_Perc2_PFall = "" Or TB_Perc2_PFall = 0) Then x(1) = TB_Perc2_PFall Else x(1) = 0
    If Not (TB_Perc3_PFall = "" Or TB_Perc3_PFall = 0 Or ChkB_Perc3.Value = False) Then x(2) = TB_Perc3_PFall Else x(2) = 0
    If Not (TB_Perc4_PFall = "" Or TB_Perc4_PFall = 0 Or ChkB_Perc4.Value = False) Then x(3) = TB_Perc4_PFall Else x(3) = 0
    If Not (TB_Perc5_PFall = "" Or TB_Perc5_PFall = 0 Or ChkB_Perc5.Value = False) Then x(4) = TB_Perc5_PFall Else x(4) = 0
    If Not (TB_Perc6_PFall = "" Or TB_Perc6_PFall = 0 Or ChkB_Perc6.Value = False) Then x(5) = TB_Perc6_PFall Else x(5) = 0
    
    max = x(0)
    For i = 0 To 5
        If x(i) > max Then max = x(i)
    Next i
    
    TB_PFall_Max = max
End Sub

'------------------------------------------��������� ���������� ���������� ��������--------------------------------
Private Function TotalWorkTimeCalculate(ai_minPressure As Integer, ai_ReductorPressure As Integer, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'��������� ���������� ������ ������� ������ (� ��������)
Dim vd_Temp As Double
    '��������� ������������� �������� ������� ������ � ��������
    vd_Temp = (ai_minPressure - ai_ReductorPressure) * as_BalloonValue / ((ai_AirExpence / 60) * as_CempressFaxtor)
    '���������� ���������� ��������
    TotalWorkTimeCalculate = CLng(vd_Temp)
End Function

Private Function TimeAtFireCalculate(ai_minPressure As Integer, as_BackResrve As Single, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'��������� ���������� ������������ ������� ������ (� ��������)
Dim vd_Temp As Double
    '��������� ������������� �������� ������� ������ � ��������
    vd_Temp = (ai_minPressure - as_BackResrve) * as_BalloonValue / ((ai_AirExpence / 60) * as_CempressFaxtor)
    '���������� ���������� ��������
    TimeAtFireCalculate = CLng(vd_Temp)
End Function

Private Function fs_WorkTimeCalculateUntilFireFind(ai_maxFallPressure As Integer, _
                                        as_BalloonValue As Single, ai_AirExpence As Integer, _
                                        as_CempressFaxtor As Single) As Long
'��������� ���������� ������� ������ � ������� ��������� �� ������� ������ ������� �������� ��� ������������� ����� (� ��������)
Dim vd_Temp As Double
    '��������� ������������� �������� ������� ������ � ��������
    vd_Temp = (ai_maxFallPressure * as_BalloonValue) / ((ai_AirExpence / 60) * as_CempressFaxtor)
    '���������� ���������� ��������
    fs_WorkTimeCalculateUntilFireFind = CLng(vd_Temp)
End Function

Private Function fs_BackResrveCalculate(ai_maxFallPressure As Integer, ab_HardFactor As Boolean, _
                                        ai_ReductorPressure As Integer) As Single
'��������� ���������� ������ �������� ������������ ��� �����������, ���
    If ab_HardFactor = True Then
        fs_BackResrveCalculate = ai_maxFallPressure * 2 + ai_ReductorPressure
    Else
        fs_BackResrveCalculate = ai_maxFallPressure * 1.5 + ai_ReductorPressure
    End If
End Function

Private Function fs_MaxPressureFallWF(ai_minEnterPressure As Integer, ab_HardFactor As Boolean, _
                                        ai_ReductorPressure As Integer) As Single
'��������� ���������� ����������� ���������� ������� �������� � ���������, ���
    If ab_HardFactor = True Then
        fs_MaxPressureFallWF = (ai_minEnterPressure - ai_ReductorPressure) / 3
    Else
        fs_MaxPressureFallWF = (ai_minEnterPressure - ai_ReductorPressure) / 2.5
    End If
End Function

Private Function fB_ConvHardFactor(aS_HardFactor As String) As Boolean
'������� ���������� ������, ���� ������� ������� � ����, ���� �����������
    If aS_HardFactor = "����������� �������" Then
        fB_ConvHardFactor = False
    Else
        fB_ConvHardFactor = True
    End If
End Function

Private Sub ps_ResultsRecalc()
'��������� ��������� ����������� �������
Dim vDt_Time As Date
    
    Min_P1
    Min_P2
    Max_PFall
    
    On Error GoTo EX
    
    '---��������� ����� ������ ��� �� ������������ �����
    TB_MaxFall = _
        CStr(fs_MaxPressureFallWF(CInt(TB_P1_Min), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))

    TB_ExitPressure = _
        CStr(CInt(TB_P1_Min) - fs_MaxPressureFallWF(CInt(TB_P1_Min), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))
        
    TB_TimuUntilOrder = _
        CStr(Round(fs_WorkTimeCalculateUntilFireFind(CInt(TB_MaxFall), _
        CSng(TB_BallonsValue), CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    
    '---��������� ����� ������ ��� ������������ �����
    TB_BackWayReserv = _
        CStr(fs_BackResrveCalculate(CInt(TB_PFall_Max), _
        fB_ConvHardFactor(CB_Conditions), CInt(TB_ReductorNeedPressure)))
    TB_TimeAtFire.Value = _
        CStr(Round(TimeAtFireCalculate(CInt(TB_P2_Min), CSng(TB_BackWayReserv), CSng(TB_BallonsValue), _
        CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    TB_TimeTotal.Value = _
        CStr(Round(TotalWorkTimeCalculate(CInt(TB_P1_Min), CInt(TB_ReductorNeedPressure), CSng(TB_BallonsValue), _
        CInt(TB_DirectExpense), CSng(TB_CompFactor)) / 60, 2))
    On Error Resume Next
    '---��������� ����� ������ ������� � ������, ���� ���� �� ����� ���������
        vDt_Time = _
            DateAdd("s", CDbl(TB_TimuUntilOrder.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
        Me.TB_OrderTime.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
    '---��������� ����� ������ ������� � ������, ���� ���� ���������
        vDt_Time = _
            DateAdd("s", CDbl(Me.TB_TimeAtFire.Value) * 60, TimeValue(Me.TB_TimeArrival))
        Me.TB_OrderTimeAtFire.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)
    '---��������� ����� ������
        vDt_Time = _
            DateAdd("s", CDbl(TB_TimeTotal.Value) * 60, TimeValue(Me.TB_MainTimeEnter))
        Me.TB_MainTimeExit.Value = _
            IIf(Hour(vDt_Time) < 10, "0", "") & Hour(vDt_Time) & ":" _
            & IIf(Minute(vDt_Time) < 10, "0", "") & Minute(vDt_Time) & ":" _
            & IIf(Second(vDt_Time) < 10, "0", "") & Second(vDt_Time)

Exit Sub
EX:
    SaveLog Err, "ps_ResultsRecalc"
End Sub

'-------------------------------��������� ����������� ������ � �������������� ������--------------------------------
Private Sub ps_ValuesBack()
'��������� ����������� ������ � ������-���������
Dim vO_IntShape As Visio.Shape
Dim vD_TimeTemp As Date
Dim vL_Interval As Long
Dim vdbl_TimeTemp As Double

    On Error GoTo EX

Set vO_IntShape = Application.ActivePage.Shapes.ItemFromID(VL_InitShapeID)

'---���������� �������� ������� ����� .Prop
'    ChainContentForm.VS_DevceType = aS_SIZODType
    vO_IntShape.Cells("Prop.AirDevice").FormulaU = """" & Me.LB_DeviceModel.Column(0, Me.LB_DeviceModel.Value) & """"
'    ChainContentForm.VL_InitShapeID = ShpObj.ID
    vO_IntShape.Cells("Prop.Personnel").FormulaU = 2 + IIf(Me.ChkB_Perc3.Value, 1, 0) + IIf(Me.ChkB_Perc4.Value, 1, 0) + _
        IIf(Me.ChkB_Perc5.Value, 1, 0) + IIf(Me.ChkB_Perc6.Value, 1, 0)
    vO_IntShape.Cells("Prop.WorkPlace").FormulaU = """" & Me.CB_Conditions & """"
    vO_IntShape.Cells("Prop.AirConsuption").FormulaU = CStr(Me.TB_DirectExpense)
    vO_IntShape.Cells("Actions.ResultShow.Checked").FormulaU = Me.ChkB_ShowResults.Value
    If Me.VB_TimeChange = True Then '��������� ������ �� ������� ���������, � ������, ���� ��� ��������
        vD_TimeTemp = DateValue(vO_IntShape.Cells("Prop.FormingTime").ResultStr(visDate))
        vL_Interval = Hour(Me.TB_MainTimeEnter.Value) * 3600 + _
            Minute(Me.TB_MainTimeEnter.Value) * 60 + Second(Me.TB_MainTimeEnter.Value)
        vdbl_TimeTemp = CDbl(DateAdd("s", vL_Interval, vD_TimeTemp))
        vO_IntShape.Cells("Prop.FormingTime").FormulaU = "DATETIME(" & str(vdbl_TimeTemp) & ")"
    End If
    If Me.VB_TimeArrivalChange = True And vO_IntShape.CellExists("Prop.ArrivalTime", 0) = True Then '��������� ������ �� ������� �������� � �����, � ������, ���� ��� �������� � ������� ��������������� ������
        vD_TimeTemp = DateValue(vO_IntShape.Cells("Prop.ArrivalTime").ResultStr(visDate))
        vL_Interval = Hour(Me.TB_TimeArrival.Value) * 3600 + _
            Minute(Me.TB_TimeArrival.Value) * 60 + Second(Me.TB_TimeArrival.Value)
        vdbl_TimeTemp = CDbl(DateAdd("s", vL_Interval, vD_TimeTemp))
        vO_IntShape.Cells("Prop.ArrivalTime").FormulaU = "DATETIME(" & str(vdbl_TimeTemp) & ")"
    End If

    '---���������� ������ ��� ����������������� �1
    vO_IntShape.Cells("Scratch.A1").FormulaU = """" & Me.TB_Perc1.Value & """"
    vO_IntShape.Cells("Scratch.B1").FormulaU = Me.TB_Perc1_P1.Value
    vO_IntShape.Cells("Scratch.C1").FormulaU = Me.TB_Perc1_P2.Value
    '---������������ ������ ��� ����������������� �2
    If Me.ChkB_Perc2.Value Then
        vO_IntShape.Cells("Scratch.A2").FormulaU = """" & Me.TB_Perc2.Value & """"
        vO_IntShape.Cells("Scratch.B2").FormulaU = Me.TB_Perc2_P1.Value
        vO_IntShape.Cells("Scratch.C2").FormulaU = Me.TB_Perc2_P2.Value
    End If
    '---������������ ������ ��� ����������������� �3
    If Me.ChkB_Perc3.Value Then
        vO_IntShape.Cells("Scratch.A3").FormulaU = """" & Me.TB_Perc3.Value & """"
        vO_IntShape.Cells("Scratch.B3").FormulaU = Me.TB_Perc3_P1.Value
        vO_IntShape.Cells("Scratch.C3").FormulaU = Me.TB_Perc3_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B3").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C3").FormulaU = """" & """"
    End If

    '---������������ ������ ��� ����������������� �4
    If Me.ChkB_Perc4.Value Then
        vO_IntShape.Cells("Scratch.A4").FormulaU = """" & Me.TB_Perc4.Value & """"
        vO_IntShape.Cells("Scratch.B4").FormulaU = Me.TB_Perc4_P1.Value
        vO_IntShape.Cells("Scratch.C4").FormulaU = Me.TB_Perc4_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B4").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C4").FormulaU = """" & """"
    End If
    '---������������ ������ ��� ����������������� �5
    If Me.ChkB_Perc5.Value Then
        vO_IntShape.Cells("Scratch.A5").FormulaU = """" & Me.TB_Perc5.Value & """"
        vO_IntShape.Cells("Scratch.B5").FormulaU = Me.TB_Perc5_P1.Value
        vO_IntShape.Cells("Scratch.C5").FormulaU = Me.TB_Perc5_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B5").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C5").FormulaU = """" & """"
    End If
    '---������������ ������ ��� ����������������� �6
    If Me.ChkB_Perc6.Value Then
        vO_IntShape.Cells("Scratch.A6").FormulaU = """" & Me.TB_Perc6.Value & """"
        vO_IntShape.Cells("Scratch.B6").FormulaU = Me.TB_Perc6_P1.Value
        vO_IntShape.Cells("Scratch.C6").FormulaU = Me.TB_Perc6_P2.Value
    Else
        vO_IntShape.Cells("Scratch.B6").FormulaU = """" & """"
        vO_IntShape.Cells("Scratch.C6").FormulaU = """" & """"
    End If

Exit Sub
EX:
    SaveLog Err, "ps_ValuesBack"
End Sub





