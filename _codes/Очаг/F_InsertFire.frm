VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InsertFire 
   Caption         =   "�������� ������"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   OleObjectBlob   =   "F_InsertFire.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InsertFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Vfl_TargetShapeID As Long
Public VmD_TimeStart As Date
Public VmD_Time1Stvol As Date
'Private VfB_NotShowPropertiesWindow As Boolean      '�������� ����� �������
Private vfStr_ObjList() As String


Dim matrixSize As Long              '���������� ������ � �������
Dim matrixChecked As Long           '���������� ����������� ������
Public timeElapsedMain As Single    '����� ��������� � ������ �������������
Public pathMain As Single           '���������� ���� � ������ �������������


'--------------------------------���� ����������---------------------------------------------------
Private Sub B_Cancel_Click()
'---�������� ����� ����
    VfB_NotShowPropertiesWindow = False
    
    Me.Hide
End Sub




Private Sub B_OK_Click()

'---��������� ������������ ��������� ������������� ������
    If fC_DataCheck = False Then Exit Sub

    Me.Hide
    
'---��������� ����� ����� ���������� ������� ������� � � ����������� �� �����
    '���������� ������� ����� ������� ��� ������������
    If Me.OB_SolidShape.value = True Then
        s_FireShapeDrop
    ElseIf Me.OB_LiquidShape.value = True = True Then
        s_PrognoseFire
    End If
    
'---�������� ����� ����
    VfB_NotShowPropertiesWindow = False
End Sub

Private Sub B_Test_Click()
    s_PrognoseFire
End Sub






'--------------------------------���� �������������---------------------------------------------------
Private Sub B_Cancel2_Click()
'---�������� ����� ����
    VfB_NotShowPropertiesWindow = False
    
    Me.Hide
End Sub

Private Sub btnBakeMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "������� ������� ������ �������������� �������! ��������� ������ �������� ����� ��� ����� �������.", vbInformation, "������-������"
        Exit Sub
    End If

    '���������� �������� ����� �������
    grain = Me.txtGrainSize
    
    '������� ������ ���� Fire
    ClearLayer "Fire"
    
    '�������� �������
    MakeMatrix Me
End Sub

Private Sub btnDeleteMatrix_Click()
    '������� �������
    DestroyMatrix
    
    '������� ������ ���� Fire
    ClearLayer "Fire"
    
    '���������, ��� ������� �� ��������
    lblMatrixIsBaked.Caption = "������� �� ��������."
    lblMatrixIsBaked.ForeColor = vbRed
End Sub

Private Sub btnRefreshMatrix_Click()
    If IsAcceptableMatrixSize(1200000, Me.txtGrainSize.value) = False Then
        MsgBox "������� ������� ������ �������������� �������! ��������� ������ �������� ����� ��� ����� �������.", vbInformation, "������-������"
        Exit Sub
    End If
    
    '��������� ������� �������� �����������
    RefreshOpenSpacesMatrix Me
End Sub

Private Sub btnRunFireModelling_Click()
'��� ������� �� ������ ��������� �������������
    '���������, �������� �� �������
    If Not IsMatrixBacked Then
        MsgBox "������� �� ��������!!!"
        Exit Sub
    End If
    
    stopModellingFlag = False
    
    On Error GoTo ex
    '���������� ��������� ���������� �����
    Dim spd As Single
    Dim timeElapsed As Single
    Dim intenseNeed As Single
    
    '���������� �������� ��������
    spd = GetSpeed
    '���������� ����� �������������
    '---���������� ����� ������������� � ������������ � ���������� ������������� ����������
    Dim vsStr_FirePath As String
    Dim vsD_TimeCur As Date
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        timeElapsed = DateDiff("s", VmD_TimeStart, vsD_TimeCur) / 60 - timeElapsedMain ' � �������!!!!!!
    End If
    If Me.OB_ByDuration = True Then
        timeElapsed = Me.TB_Duration
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = str(ffSng_PointChange(Me.TB_Radius) * 2) & "m"
        If timeElapsedMain > 10 Then
            timeElapsed = CSng(ffSng_PointChange(Me.TB_Radius) / spd)      ' � �������!!!!!!
        Else
            timeElapsed = CSng(ffSng_PointChange(Me.TB_Radius) / (spd / 2))    ' � �������!!!!!!
        End If
    End If
    
    '���������� ��������� �������������
    intenseNeed = GetIntense
    
    '���������, ��� �� ������ ������� �����
    If timeElapsed > 0 And spd > 0 Then
        '������ �������
        If Me.OB_ByRadius = True Then       '���� ������������� ��������������� �� �������...
            RunFire timeElapsed, spd, intenseNeed, CSng(ffSng_PointChange(Me.TB_Radius))
        Else
            RunFire timeElapsed, spd, intenseNeed
        End If
        
    Else
        MsgBox "�� ��� ������ ��������� �������!", vbCritical
        Exit Sub
    End If
    
    '---�������� ���������� ������ �������� ������� ������ �� ����� ���������� �������
    Dim vsO_FireShape As Visio.Shape
    Set vsO_FireShape = Application.ActiveWindow.Selection(1)
    If Me.OB_SpeedByObject.value = True Then
        vsO_FireShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_FireShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_FireShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
    End If
    '---��������� ����������� ����/����� ��� ���������� ������ ������� �������
    Dim actTime As Date
    actTime = DateAdd("n", timeElapsedMain, VmD_TimeStart)
    vsO_FireShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(actTime)) & ")"
        
Exit Sub
ex:
    MsgBox "�� ��� ������ ��������� �������!", vbCritical
End Sub

Private Sub btnStopModelling_Click()
'������ ����� � �������������
    stopModellingFlag = True
End Sub


Private Sub optTTX_Change()
'    txtNozzleRangeValue.Enabled = False
End Sub
Private Sub optValue_Change()
    txtNozzleRangeValue.Enabled = optValue.value
    If txtNozzleRangeValue.Enabled Then
        txtNozzleRangeValue.BackColor = vbWhite
    Else
        txtNozzleRangeValue.BackColor = &H8000000F
    End If
End Sub

'--------------------------��������� ��������� �������������----------------------------------
Private Function GetMatrixCheckedStatus(Optional kind As Byte = 0) As String
'���������� ������� ��� ������� ��������� �������
Dim procent As Single
    procent = Round(matrixChecked / matrixSize, 4) * 100
    
    If kind = 0 Then
        GetMatrixCheckedStatus = "�������� " & procent & "%"
    ElseIf kind = 1 Then
        GetMatrixCheckedStatus = "��������� ��������� ���� " & procent & "%"
    End If
End Function

'--------------------------������� ��������� � ������� �������������--------------------------
Public Sub SetMatrixSize(ByVal size As Long)
'��������� ��� ����� ����� ���-�� ������ � �������
    matrixSize = size
    matrixChecked = 0
End Sub

Public Sub AddCheckedSize(ByVal size As Long, Optional kind As Byte = 0)
'��������� ���-�� ����������� ������
    matrixChecked = matrixChecked + size
    
    '��������� ��������� ������ � ����������� ����������� ������
    lblMatrixIsBaked.Caption = GetMatrixCheckedStatus(kind)
    lblMatrixIsBaked.ForeColor = vbBlack
'    Me.Repaint
End Sub

Public Sub Refresh()
    Me.Repaint
End Sub





'--------------------------------���� ������ ���� �������� ����---------------------------------------------------
Private Sub OB_ByDuration_AfterUpdate()
    Me.TB_Duration.Enabled = True
    Me.TB_Radius.Enabled = False
    Me.TB_Time.Enabled = False
    Me.TB_Duration.BackStyle = fmBackStyleOpaque
    Me.TB_Radius.BackStyle = fmBackStyleTransparent
    Me.TB_Time.BackStyle = fmBackStyleTransparent
    
    
End Sub


Private Sub OB_ByRadius_AfterUpdate()
    Me.TB_Radius.Enabled = True
    Me.TB_Duration.Enabled = False
    Me.TB_Time.Enabled = False
    Me.TB_Radius.BackStyle = fmBackStyleOpaque
    Me.TB_Duration.BackStyle = fmBackStyleTransparent
    Me.TB_Time.BackStyle = fmBackStyleTransparent
    
    
End Sub


Private Sub OB_ByTime_AfterUpdate()
    Me.TB_Time.Enabled = True
    Me.TB_Duration.Enabled = False
    Me.TB_Radius.Enabled = False
    Me.TB_Time.BackStyle = fmBackStyleOpaque
    Me.TB_Duration.BackStyle = fmBackStyleTransparent
    Me.TB_Radius.BackStyle = fmBackStyleTransparent
    
    
End Sub


'--------------------------------���� ������ ��������----------------------------------------------------------------
Private Sub OB_SpeedByObject_AfterUpdate()
    Me.CB_Object.Enabled = True
    Me.CB_ObjectType.Enabled = True
    Me.TB_Speed2.Enabled = False
    Me.TB_Intense2.Enabled = False
    Me.CB_Object.BackStyle = fmBackStyleOpaque
    Me.CB_ObjectType.BackStyle = fmBackStyleOpaque
    Me.TB_Speed2.BackStyle = fmBackStyleTransparent
    Me.TB_Speed1.BackStyle = fmBackStyleOpaque
    Me.TB_Intense2.BackStyle = fmBackStyleTransparent
    Me.TB_Intense1.BackStyle = fmBackStyleOpaque

End Sub


Private Sub OB_SpeedByDirect_AfterUpdate()
    Me.CB_Object.Enabled = False
    Me.CB_ObjectType.Enabled = False
    Me.TB_Speed2.Enabled = True
    Me.TB_Intense2.Enabled = True
    Me.CB_Object.BackStyle = fmBackStyleTransparent
    Me.CB_ObjectType.BackStyle = fmBackStyleTransparent
    Me.TB_Speed2.BackStyle = fmBackStyleOpaque
    Me.TB_Speed1.BackStyle = fmBackStyleTransparent
    Me.TB_Intense2.BackStyle = fmBackStyleOpaque
    Me.TB_Intense1.BackStyle = fmBackStyleTransparent

End Sub


'--------------------------------���� ������ ���� ����������-----------------------------------------------------------
Private Sub OB_SolidShape_AfterUpdate()
    Me.TB_REI.Enabled = False
    Me.TB_REI.BackStyle = fmBackStyleTransparent
    Me.CB_CheckOpens.Enabled = False
    Me.CB_CheckOpens.BackStyle = fmBackStyleTransparent
    
    Me.CB_Shape.Enabled = True
    Me.CB_Shape.BackStyle = fmBackStyleOpaque
End Sub

Private Sub OB_LiquidShape_AfterUpdate()
    Me.TB_REI.Enabled = True
    Me.TB_REI.BackStyle = fmBackStyleOpaque
    Me.CB_CheckOpens.Enabled = True
    Me.CB_CheckOpens.BackStyle = fmBackStyleOpaque
    
    Me.CB_Shape.Enabled = False
    Me.CB_Shape.BackStyle = fmBackStyleTransparent
End Sub



'--------------------------------���� �������� ��������----------------------------------------------------------------

Private Sub UserForm_Initialize()
'��������� �������� �����

'---��������� ������:
    sf_ObjectsListCreation '������� ������ �������� � ���������
    
    '---�������� ���������� �������
    sf_ObjectTypesListRefresh
    sf_ObjectsListRefresh
    sf_FireFormLoad
    sf_StvolsOptionsLoad
    
End Sub

Public Property Get AttackDeep() As Integer
    AttackDeep = 5000   '������� ������� ������� ��������� �� ��������� ������ 5�
End Property
Public Property Get StvolCalcDistance() As Integer
    StvolCalcDistance = txtNozzleRangeValue * 1000
End Property

Private Sub UserForm_Activate()
'��������� ��������� ����� - ��� ������
'---��������� ����� ����
    VfB_NotShowPropertiesWindow = True

'---��������� �������� ������� � �������� � �������
'    Me.TB_Radius.Value = ffSng_PointChange(Me.TB_Radius.Value)  ����������� �����������!!!!

'---���������� ��������� �������������
    matrixSize = 0
    matrixChecked = 0
    
    '���������, �������� �� �������
    If IsMatrixBacked Then
        lblMatrixIsBaked.Caption = "������� ��������. ������ ����� " & grain & "��."
        lblMatrixIsBaked.ForeColor = vbGreen
        Me.txtGrainSize = grain
    Else
        lblMatrixIsBaked.Caption = "������� �� ��������."
        lblMatrixIsBaked.ForeColor = vbRed
        Me.txtGrainSize.value = 200
    End If
    
    
'    VmD_TimeStart = Application.act
    
End Sub

Private Sub sf_ObjectTypesListRefresh()
'��������� ���������� ������ ��������� ��������
Dim vsO_DBS As Object, vsO_RST As Object
Dim vsStr_SQL As String
Dim vsStr_Pth As String

    On Error GoTo ex
'---������� ��������� ������
'    If CB_ObjectType.ListCount > 0 Then Exit Sub '� ������, ���� ������ ��� �������� - �� ��������� ���
    Me.CB_ObjectType.Clear
'    Me.CB_Object.Clear

'---���������� ������ SQL ��� ������ ������� ��������� �� ���� ������
    vsStr_SQL = "SELECT �����������������.��������� FROM �����������������;"
    
'---������� ����� ������� ��� ��������� ������ ���������
    vsStr_Pth = ThisDocument.path & "Signs.fdb"
'    Set vsO_DBS = GetDBEngine.OpenDatabase(vsStr_Pth)
'    Set vsO_RST = vsO_DBS.CreateQueryDef("", vsStr_SQL).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
    Set vsO_DBS = CreateObject("ADODB.Connection")
    vsO_DBS = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & vsStr_Pth & ";Uid=Admin;Pwd=;"
    vsO_DBS.Open
    Set vsO_RST = CreateObject("ADODB.Recordset")
    vsO_RST.Open vsStr_SQL, vsO_DBS, 3, 1
'    Set RSField = vsO_RST.Fields(FieldName)

'---���� ����������� ������ � ������ ������ � ��������� � � ������ ��������� ��������
    With vsO_RST
        .MoveFirst
        Do Until .EOF
            Me.CB_ObjectType.AddItem !���������
            .MoveNext
        Loop
    End With

'---���������� ������ ������� ������
    Me.CB_ObjectType.ListIndex = 0

'---������� �������
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
Exit Sub
ex:
'---������� �������
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
    SaveLog Err, "sf_ObjectTypesListRefresh"
End Sub


Private Sub sf_ObjectsListCreation()
'��������� ��������� ������ �������� ������
Dim vsO_DBS As Object, vsO_RST As Object
Dim vsStr_SQL As String
Dim vsStr_Pth As String
Dim i As Integer

    On Error GoTo ex
'---���������� ������ SQL ��� ������ ������� ��������� �� ���� ������
    vsStr_SQL = "SELECT ���������, ��������, ������������, ����������������������� FROM �_�������������;"
    
'---������� ����� ������� ��� ��������� ������ ���������
    vsStr_Pth = ThisDocument.path & "Signs.fdb"
'    Set vsO_DBS = GetDBEngine.OpenDatabase(vsStr_Pth)
'    Set vsO_RST = vsO_DBS.CreateQueryDef("", vsStr_SQL).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
    Set vsO_DBS = CreateObject("ADODB.Connection")
    vsO_DBS = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & vsStr_Pth & ";Uid=Admin;Pwd=;"
    vsO_DBS.Open
    Set vsO_RST = CreateObject("ADODB.Recordset")
    vsO_RST.Open vsStr_SQL, vsO_DBS, 3, 1
'    Set RSField = vsO_RST.Fields(FieldName)

'---���� ����������� ������ � ������ ������ � ��������� � � ������ ��������� ��������
    With vsO_RST
    i = 0
    '---������������� ��������� ������
        vsO_RST.MoveLast
        ReDim vfStr_ObjList(vsO_RST.RecordCount, 4) As String
    '---��������� ������ ��������� ������ � �� ���������
        .MoveFirst
        Do Until .EOF
            vfStr_ObjList(i, 0) = !���������
            vfStr_ObjList(i, 1) = !��������
            If !������������ >= 0 Then vfStr_ObjList(i, 2) = !������������ Else vfStr_ObjList(i, 2) = 0                             '�������� ���������������, ��������
            If !����������������������� >= 0 Then vfStr_ObjList(i, 3) = !����������������������� Else vfStr_ObjList(i, 3) = "0.1"   '������������� �� ����
            i = i + 1
            .MoveNext
        Loop
    End With

'---������� �������
    Set vsO_DBS = Nothing
    Set vsO_RST = Nothing
Exit Sub
ex:
    SaveLog Err, "sf_ObjectsListCreation"
End Sub

Private Sub sf_ObjectsListRefresh()
'��������� ���������� ������ �������� ������
Dim i As Integer

Me.CB_Object.Clear

For i = 0 To UBound(vfStr_ObjList()) - 1
    If vfStr_ObjList(i, 0) = Me.CB_ObjectType.value Then
        Me.CB_Object.AddItem vfStr_ObjList(i, 1)
    End If
Next i

Me.CB_Object.ListIndex = 0

End Sub




Private Sub CB_ObjectType_Change()
'��������� ������ �������� � ������������ � ����� ��������� ���������
    sf_ObjectsListRefresh
End Sub

Private Sub CB_Object_Change()
'�������� �������� �������� � ����������� �� ���������� ������� ������
Dim i As Integer

    For i = 0 To UBound(vfStr_ObjList()) - 1
        If vfStr_ObjList(i, 1) = Me.CB_Object.value Then
            Me.TB_Speed1.value = vfStr_ObjList(i, 2)
            Me.TB_Intense1.value = vfStr_ObjList(i, 3)
        End If
    Next i

End Sub

Private Sub sf_FireFormLoad()
'���������� ������ ���� ������

    With Me.CB_Shape
        .AddItem "������� �������������"
        .AddItem "������� �������"
        .AddItem "������ 90"
        .AddItem "������ 180"
        .AddItem "������ 270"
        .ListIndex = 0
    End With

End Sub

Private Sub sf_StvolsOptionsLoad()
'���������� ������ ����� ����� �������
    With Me.CB_StvolsOptions
        .AddItem "���� �������"
        .AddItem "������� ��������� ������"
        .ListIndex = 0
    End With
End Sub


'-------------------------------------------���� ������ ������ ������� �������---------------------------------------

Private Sub s_FireShapeDrop()
'��������� ������ ������ ������� � ����������� � ��������� ����������
Dim vsO_DropMaster As Visio.Master
Dim vsO_DropShape As Visio.Shape
Dim vsO_DropTargetShape As Visio.Shape
Dim vss_Speed As Single
Dim vsStr_FirePath As String
Dim vsD_TimeCur As Date
Dim vsL_Duration As Long

'---���������� ��������� ������
    On Error GoTo ex
'---���������� ������� �������
    Set vsO_DropTargetShape = Application.ActivePage.Shapes.ItemFromID(Vfl_TargetShapeID)
    Set vsO_DropMaster = ThisDocument.Masters(Me.CB_Shape.value)

'---���������� ������ ������� ������� � ���������� � �����������������
    Set vsO_DropShape = Application.ActivePage.Drop(vsO_DropMaster, 0, 0)
    vsO_DropShape.Cells("PinX").FormulaU = vsO_DropTargetShape.Cells("PinX").FormulaU
    vsO_DropShape.Cells("PinY").FormulaU = vsO_DropTargetShape.Cells("PinY").FormulaU
    vsO_DropShape.SendToBack

'---���������� ���������� �������� ��������������� ����
    vss_Speed = GetSpeed
'    If Me.OB_SpeedByDirect = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed2))
'    ElseIf Me.OB_SpeedByObject = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed1))
'    End If
    
'---���������� ������� ������ � ������������ � ����� ���������� ����� � ������������ _
    � ���������� ������������� ����������
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        vsL_Duration = DateDiff("s", VmD_TimeStart, vsD_TimeCur) ' � ��������!!!!!!
        If vsL_Duration / 60 > 10 Then
            vsStr_FirePath = str((vsL_Duration - 300) * (vss_Speed / 60) * 2) & "m"
        Else
            vsStr_FirePath = str(vsL_Duration * (vss_Speed / 60)) & "m"
        End If
    End If
    If Me.OB_ByDuration = True Then
        If ffSng_PointChange(Me.TB_Duration) > 10 Then
            vsStr_FirePath = str((ffSng_PointChange(Me.TB_Duration) - 5) * vss_Speed * 2) & "m"
        Else
            vsStr_FirePath = str(ffSng_PointChange(Me.TB_Duration) * vss_Speed) & "m"
        End If
        vsD_TimeCur = DateAdd("n", ffSng_PointChange(Me.TB_Duration), VmD_TimeStart)
        vsD_TimeCur = DateAdd("s", Ost(ffSng_PointChange(Me.TB_Duration)) * 60, vsD_TimeCur)
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = str(ffSng_PointChange(Me.TB_Radius) * 2) & "m"
        If ffSng_PointChange(Me.TB_Radius) / (vss_Speed * 0.5) > 10 Then
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / vss_Speed + 5) * 60      ' � ��������!!!!!!
        Else
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / (vss_Speed / 2)) * 60    ' � ��������!!!!!!
        End If

        vsD_TimeCur = DateAdd("s", vsL_Duration, VmD_TimeStart)
        
    End If
    '---������ ����������� ������� � ��������� ����������� ����
    If vsO_DropMaster.Name = "������ 90" Then
        vsStr_FirePath = str(ffSng_PointChange(val(vsStr_FirePath) / 2)) & "m"
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    ElseIf vsO_DropMaster.Name = "������ 180" Then
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsStr_FirePath = str(ffSng_PointChange(val(vsStr_FirePath) / 2)) & "m"
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    Else
        vsO_DropShape.Cells("Width").FormulaU = vsStr_FirePath
        vsO_DropShape.Cells("Height").FormulaU = vsStr_FirePath
    End If
    vsO_DropShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(vsD_TimeCur)) & ")"
    
'---�������� ���������� ������ �������� ������� ������ �� ����� ���������� �������
    If Me.OB_SpeedByObject.value = True Then
        vsO_DropShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_DropShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_DropShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
'        vsO_DropShape.Cells("Prop.FireSpeedLine").FormulaU = CDbl(Me.TB_Speed2.Value)
    End If
       
    
'---������� �������
    Set vsO_DropTargetShape = Nothing
    Set vsO_DropShape = Nothing
    Set vsO_DropMaster = Nothing
    
Exit Sub

ex:
'MsgBox "���� �� ��������� ���� �������� ������� ������ ��� ������� � ��������! " & _
'    "��������� ������������ ��������� ������!", vbCritical
    MsgBox "� �������� ������ ��������� �������� ������! ��������� � ������������ ��������� ���� ������.", , ThisDocument.Name
'---������� �������
    Set vsO_DropTargetShape = Nothing
    Set vsO_DropShape = Nothing
    Set vsO_DropMaster = Nothing
    SaveLog Err, "s_FireShapeDrop"
End Sub


Private Sub s_PrognoseFire()
'��������� ������ ������� �������� ������, �������� ��� � ���� ������� � ������� ���������� ������ ������������ ���������
Dim vO_Fire As c_Fire
Dim x As Double, y As Double
Dim shp As Visio.Shape
Dim vsO_FireShape As Visio.Shape
Dim vss_Speed As Single '�������� ��������������� ����
Dim vsStr_FirePath As String '���� ���������� �����
Dim vsD_TimeCur As Date
Dim vsL_Duration As Long

On Error GoTo Tail

'---������� ��������� ������ c_Fire ��� ��������������� ������� �������
    Set vO_Fire = New c_Fire

'---���������� ������ ����� ������ � ��������� ����������
    Set shp = Application.ActiveWindow.Selection(1)
    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)

'---���������� ���������� �������� ��������������� ����
    vss_Speed = GetSpeed
'    If Me.OB_SpeedByDirect = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed2))
'    ElseIf Me.OB_SpeedByObject = True Then
'        vss_Speed = CSng(ffSng_PointChange(Me.TB_Speed1))
'    End If

'---���������� ������� ������ � ������������ � ����� ���������� ����� � ������������ _
    � ���������� ������������� ����������
    If Me.OB_ByTime = True Then
        vsD_TimeCur = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
        vsL_Duration = DateDiff("s", VmD_TimeStart, vsD_TimeCur) ' � ��������!!!!!!
        If vsL_Duration / 60 > 10 Then
            vsStr_FirePath = (vsL_Duration - 300) * (vss_Speed / 60)
        Else
            vsStr_FirePath = (vsL_Duration * (vss_Speed / 60)) / 2
        End If
    End If
    If Me.OB_ByDuration = True Then
        If ffSng_PointChange(Me.TB_Duration) > 10 Then
            vsStr_FirePath = (ffSng_PointChange(Me.TB_Duration) - 5) * vss_Speed
        Else
            vsStr_FirePath = (ffSng_PointChange(Me.TB_Duration) * vss_Speed) / 2
        End If
        vsD_TimeCur = DateAdd("n", ffSng_PointChange(Me.TB_Duration), VmD_TimeStart)
        vsD_TimeCur = DateAdd("s", Ost(ffSng_PointChange(Me.TB_Duration)) * 60, vsD_TimeCur)
    End If
    If Me.OB_ByRadius = True Then
        vsStr_FirePath = ffSng_PointChange(Me.TB_Radius)
        If ffSng_PointChange(Me.TB_Radius) / (vss_Speed * 0.5) > 10 Then
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / vss_Speed + 5) * 60      ' � ��������!!!!!!
        Else
            vsL_Duration = CSng(ffSng_PointChange(Me.TB_Radius) / (vss_Speed / 2)) * 60    ' � ��������!!!!!!
        End If

        vsD_TimeCur = DateAdd("s", vsL_Duration, VmD_TimeStart)
        
    End If

'---��������������� ������ ������� �������� ������
    With vO_Fire
        .PB_CheckOpens = Me.CB_CheckOpens.value
        .PI_DoorsREI = CInt(ffSng_PointChange(Me.TB_REI))
        .PS_LineSpeedM = vss_Speed
        .S_SetFullShape x, y, vsStr_FirePath
    End With

'---���������� ������������ ������ � �������� �� � ������ ������� �������
    Set vsO_FireShape = Application.ActiveWindow.Selection(1)
    ImportAreaInformation
    

    
'---�������� ���������� ������ �������� ������� ������ �� ����� ���������� �������
    If Me.OB_SpeedByObject.value = True Then
        vsO_FireShape.Cells("Prop.FireCategorie").FormulaU = Chr(34) & Me.CB_ObjectType.value & Chr(34)
        vsO_FireShape.Cells("Prop.FireDescription").FormulaU = Chr(34) & Me.CB_Object.value & Chr(34)
    Else
        vsO_FireShape.Cells("Prop.FireSpeedLine").FormulaU = Chr(34) & ffSng_PointChange(Me.TB_Speed2.value) & Chr(34)
    End If
    '---��������� ����������� ����/����� ��� ���������� ������ ������� �������
    vsO_FireShape.Cells("Prop.SquareTime").FormulaU = "DateTime(" & str(CDbl(vsD_TimeCur)) & ")"


Set vsO_FireShape = Nothing
Set vO_Fire = Nothing
Set shp = Nothing
Set vO_Fire = Nothing

Exit Sub
Tail:
'    Debug.Print Err.Description
    Set vsO_FireShape = Nothing
    Set vO_Fire = Nothing
    Set shp = Nothing
    Set vO_Fire = Nothing
    SaveLog Err, "s_PrognoseFire"
End Sub













'-------------------------------------���� ���������/���������������� �������� � �������-------------------------
Private Function ffSng_PointChange(afStr_String) As Single
'������� �������������� ����� � �������
Dim i As Integer
Dim vfStr_TempString As String

For i = 1 To Len(afStr_String)
    If Mid(afStr_String, i, 1) = "." Then
        vfStr_TempString = vfStr_TempString & ","
    Else
        vfStr_TempString = vfStr_TempString & Mid(afStr_String, i, 1)
    End If
Next i

ffSng_PointChange = CSng(vfStr_TempString)


End Function

Function Ost(Count As Single) As Single
'������� ���������� ������� ���� �����, � ��������� �� �����
Ost = Round(Count - Int(Count), 2)
End Function

Function GetSpeed() As Single
'---���������� ���������� �������� ��������������� ����
    If Me.OB_SpeedByDirect = True Then
        GetSpeed = CSng(ffSng_PointChange(Me.TB_Speed2))
    ElseIf Me.OB_SpeedByObject = True Then
        GetSpeed = CSng(ffSng_PointChange(Me.TB_Speed1))
    End If
End Function

Function GetIntense() As Single
'---���������� ��������� ������������� ������ ����
    If Me.OB_SpeedByObject = True Then
        GetIntense = CSng(ffSng_PointChange(Me.TB_Intense1))
    ElseIf Me.OB_SpeedByDirect = True Then
        GetIntense = CSng(ffSng_PointChange(Me.TB_Intense2))
    End If
End Function



'--------------------------------������� �������� ������������ ��������� ������----------------------------------
Private Function fC_DataCheck() As Boolean
'������� ������� ���������� ���������� �������� ������������ ������ ��������� ������������� � �����
    fC_DataCheck = False

    If Me.OB_ByTime.value = True Then
        If fC_DateCorrCheck = False Then Exit Function
        If fC_DateDiffCheck = False Then Exit Function
    End If
    If Me.OB_ByDuration.value = True Then
        If fC_DurationCheck = False Then Exit Function
        If fC_DurationValueCheck = False Then Exit Function
    End If
    If Me.OB_ByRadius.value = True Then
'        If fC_RadiusCheck = False Then Exit Function     !!!1�������� ���������
        If fC_RadiusValueCheck = False Then Exit Function
    End If
    If Me.OB_LiquidShape.value = True Then
        If fB_REICheck = False Then Exit Function
    End If
    
    fC_DataCheck = True

End Function

Private Function fB_REICheck() As Boolean
'������� ���������� ����, ���� � ���� TB_REI �������� ������������ ������
    If IsNumeric(Me.TB_REI.value) Then
        fB_REICheck = True
    Else
        MsgBox "�������� ������� �������������, ������� �����������! ��������� ������������ ����� ������", vbCritical
        fB_REICheck = False
    End If

End Function



Private Function fC_DateDiffCheck() As Boolean
'������� ���������� ���������� �������� ������������ �������� ������ ��������� ������������� � ����� � �� �����
Dim vfD_tmpDate As Date

'---������� ��������� ����������
    vfD_tmpDate = DateValue(Me.TB_Time) + TimeValue(Me.TB_Time)
    If DateDiff("n", VmD_TimeStart, vfD_tmpDate) < 90 And Not DateDiff("n", VmD_TimeStart, vfD_tmpDate) = 0 Then
        fC_DateDiffCheck = True
    ElseIf DateDiff("n", VmD_TimeStart, vfD_tmpDate) = 0 Then
        MsgBox "������� ������� ����� ��������� ��������� ���� � ����������, �������� �����, �����������!" & _
        " ��� �������� � �������� ������� ������� �������!", vbCritical
        fC_DateDiffCheck = False
    Else
        MsgBox "������� ������� ����� ��������� ��������� ���� � ����������, �������� �����, ������� ������!" & _
        " ��� �������� � �������� ������� ������� ������� ������� � ����� ��������� �������� �� ������������������ �������!", vbCritical
        fC_DateDiffCheck = False
    End If

End Function

Private Function fC_DateCorrCheck() As Boolean
'������� ���������� ���������� �������� ������������ ����/������� ��������� ������������� � �����
Dim vfD_tmpDate As Date
On Error GoTo ErrMsg

'---�������� ������� ��������� ����������
    vfD_tmpDate = DateValue(Me.TB_Time)
    If vfD_tmpDate > -1 Then '���� ��������� ���������� �������, �� ���������� ������
        fC_DateCorrCheck = True
    End If
    Exit Function

ErrMsg:
    MsgBox "��������� ���� �������� ����/������� �� ������������� ������� - ��������� ������������ ��������� ������!", vbCritical
    fC_DateCorrCheck = False
End Function
Private Function fC_DurationCheck() As Boolean
'������� ���������� ���������� �������� ������������ ������������ �������� ������� �������� ���������� ������������� � �����

    If IsNumeric(Me.TB_Duration.value) Then
        fC_DurationCheck = True
    Else
        MsgBox "�������� ������� ��������������� ����, ������� �����������! ��������� ������������ ����� ������", vbCritical
        fC_DurationCheck = False
    End If

End Function

Private Function fC_DurationValueCheck() As Boolean
'������� ���������� ���������� �������� ������������ ���������� ������� �������� ���������� ������������� � �����

    If ffSng_PointChange(Me.TB_Duration.value) < 90 And Not ffSng_PointChange(Me.TB_Duration.value) = 0 Then
        fC_DurationValueCheck = True
    ElseIf ffSng_PointChange(Me.TB_Duration.value) = 0 Then
        MsgBox "��������� ���� �������� ������� ��������������� ����, ����� ����!" & _
        " ��� �������� � �������� ������� ������� �������!", vbCritical
        fC_DurationValueCheck = False
    Else
        MsgBox "��������� ���� �������� ������� ��������������� ����, ������� ������!" & _
        " ��� �������� � �������� ������� ������� ������� ������� � ����� ��������� �������� �� ������������������ �������!", vbCritical
        fC_DurationValueCheck = False
    End If

End Function

Private Function fC_RadiusCheck() As Boolean
'������� ���������� ���������� �������� ������������ ������� ���������� ������������� � �����

    If IsNumeric(Me.TB_Radius.value) Then
        fC_RadiusCheck = True
    Else
        MsgBox "�������� ���� ����������� �����, ������� �����������! ��������� ������������ ����� ������", vbCritical
        fC_RadiusCheck = False
    End If
End Function

Private Function fC_RadiusValueCheck() As Boolean
'������� ���������� ���������� �������� ������������ ������� ������� ���������� ������������� � �����

    If ffSng_PointChange(Me.TB_Radius.value) < 100 And Not ffSng_PointChange(Me.TB_Radius.value) = 0 Then
        fC_RadiusValueCheck = True
    ElseIf ffSng_PointChange(Me.TB_Radius.value) = 0 Then
        MsgBox "��������� ���� �������� ���� ����������� �����, ����� ����!" & _
        " ��� �������� � �������� ������� ������� �������!", vbCritical
        fC_RadiusValueCheck = False
    Else
        MsgBox "��������� ���� �������� ���� ����������� �����, ������� ������!" & _
        " ��� �������� � �������� ������� ������� ������� ������� � ����� ��������� �������� �� ������������������ �������!", vbCritical
        fC_RadiusValueCheck = False
    End If
End Function
