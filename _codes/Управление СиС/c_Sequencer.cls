VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Sequencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ����� ��� ������������� ���������� �������� �������� � ������� ������ ��� �� ����������
' ����� ������: ��� ���������� ����� ����� ������ ������ ��������� �� ���������� �������� ��������� (�������������, ��������, �����)


Public WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1

Private t As String






' ��������� ������
Private Sub Class_Initialize()
    Set app = Visio.Application
    Debug.Print ("sequencer activateed")
End Sub
Private Sub Class_Terminate()
    Set app = Nothing
    Debug.Print ("sequencer DEactivateed")
End Sub



' ��������� ����������� ����������
Private Sub app_ShapeAdded(ByVal shp As IVShape)
Dim mainShp As Visio.Shape
    On Error GoTo ex
'    DoEvents
    If IsGFSShape(shp) Then
        Set mainShp = GetGraFisShape(shp)
        If Not mainShp Is Nothing Then
            '---��������� �������
            If CheckAndAddCommand(mainShp, shp) Then
                GetShpData mainShp, shp
            End If
            
            
            
'            frm_Command.NewCommand
            
        End If
        
'        Debug.Print mainShp.Name
    End If

Exit Sub
ex:
    Debug.Print ("None")
End Sub

Private Function CheckAndAddCommand(ByRef parent As Visio.Shape, ByRef child As Visio.Shape) As Boolean
    
    CheckAndAddCommand = False
    
    '������� (+������������ ��������)
    If IsGFSShapeWithIP(parent, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 42, 34)) Then
        Select Case cellVal(child, "User.IndexPers")
            Case Is = indexpers.ipStvolRuch

                
                frm_Command.NewCommand parent, "2", "������ ����� ", False
                If frm_Command.isOk Then
                    setcellval parent, "Actions.Patrub.Checked", 1
                End If
            Case Is = indexpers.ipGDZSZvenoVozd
                '���� ������ ������ ����������� �����, �� ������ ��������� � ������������ �������
                If IsGFSShapeWithIP(parent, Array(34)) Then
                    GetShpData_0 parent, child
'                    frm_Command.NewCommand parent, "1", "��������� ���� ������������ ����", False
                Else
                    frm_Command.NewCommand parent, "1", "������������ ����� ����", False
                End If
'                frm_Command.NewCommand parent, "1", "������������ ����� ����", False
            Case Is = indexpers.ipGDZSPost
                frm_Command.NewCommand parent, "1", "��������� ���� ������������ ����", False
            Case Is = indexpers.ipKPP
                frm_Command.NewCommand parent, "1", "������������ ������ ��� ����", False
            Case Is = indexpers.ipLestn3Kol
                frm_Command.NewCommand parent, "2", "���������� ������������ ��������", False
            Case Is = indexpers.ipLestnPalka
                frm_Command.NewCommand parent, "1", "���������� ��������-�����", False
            Case Is = indexpers.ipLestnSturm
                frm_Command.NewCommand parent, "1", "���������� ��������� ��������", False
            Case Is = indexpers.ipRazvetvlenie
                frm_Command.NewCommand parent, "3", "��������� ������������� �������� �����, ���������� ������������", False
                If frm_Command.isOk Then
                    setcellval parent, "Actions.Patrub.Checked", 1
                End If
            Case Is = indexpers.ipKolonka
                frm_Command.NewCommand parent, "2", "���������� ���������� �� ��", False
                If frm_Command.isOk Then
                    setcellval parent, "Actions.WaterCollect.Checked", 1
                End If
            Case Is = indexpers.ipStab
                frm_Command.NewCommand parent, "1", "������������ ����������� ���� �������������", False
            Case Is = indexpers.ipDutyFace
'                setcellval child, "Prop.Unit", cellval(parent, "Prop.Unit")
'                setcellval child, "Prop.ArrivalTime", cellval(parent, "Prop.ArrivalTime")
            Case Else
                t = "0"
        End Select
        CheckAndAddCommand = True
'    Else
'        t = "0"
'        CheckAndAddCommand = False
    End If
    
    '������ ����
    If IsGFSShapeWithIP(parent, Array(46, 90)) Then
        Select Case cellVal(child, "user.IndexPers")
            Case Is = indexpers.ipStvolRuch
                frm_Command.NewCommand parent, "2", "������ ���� ������ ����� ", False
                If frm_Command.isOk Then
                    setcellval parent, "Actions.Patrub.Checked", 1
                End If
            Case Is = indexpers.ipGDZSPost
                frm_Command.NewCommand parent, "1", "��������� ���� ������������ ����", False
            Case Else
                t = "0"
        End Select
        CheckAndAddCommand = True
'    Else
'        t = "0"
'        CheckAndAddCommand = False
    End If
    
    
End Function

'Private Function AskTimeDiff(Optional ByVal d As String = "2") As String
''����������� � ������������ ����� ��������� �� ���������� �������
'    AskTimeDiff = InputBox("������� ������� ��������� �� ���������� ������ �������?", , d)
'End Function


' ��������� �������
Private Function GetGraFisShape(shp As Visio.Shape) As Visio.Shape
' ���� ������ �� ������� ���� �������� ������� (shp) � ���� ��� ������ �������� ������� ������, ���������� ��
Dim tmpShp As Visio.Shape
'Dim sel As Visio.Selection
'
'    Set sel = Application.ActivePage.SpatialSearch(cellval(shp, "PinX", visInches), cellval(shp, "PinY", visInches), _
'                                                    VisSpatialRelationCodes.visSpatialTouching + VisSpatialRelationCodes.visSpatialContainedIn, _
'                                                    0, _
'                                                    visSpatialFrontToBack + visSpatialIgnoreVisible)
'    For Each tmpShp In sel
'        If Not (tmpShp.ID = shp.ID) Then
'            ' if shape is grafis shape
'            If IsGFSShape(tmpShp) Then
'                ' if not fire or fog
'                If Not IsGFSShapeWithIP(shp, Array(64, 71)) Then
'                    Set GetGraFisShape = tmpShp
'                    Exit Function
'                End If
'            End If
'        End If
'    Next tmpShp
    
    
'    Dim OtherShape As Visio.Shape
    For Each tmpShp In Application.ActivePage.Shapes
'        If IsCorrectShapeForLabel(tmpShp) Then 'And OtherShape.CellExists("User.Version", 0) = True Then
        If Not (tmpShp.ID = shp.ID) Then
            If IsGFSShape(tmpShp) Then
                If tmpShp.HitTest(cellVal(shp, "PinX", visInches), cellVal(shp, "PinY", visInches), 10) >= 1 Then
                    If Not IsGFSShapeWithIP(shp, Array(64, 71)) Then
                        Set GetGraFisShape = tmpShp
    '                    Exit Function
                    End If
                End If
            End If
        End If

    Next tmpShp
    
'Set GetGraFisShape = Nothing
End Function

Private Function GetShpData(ByRef mainShp As Visio.Shape, shp As Visio.Shape) As Variant
Dim tmpDateTime0 As Double
Dim tmpDateTime As Double
Dim tmpDateTime0S As String
Dim tmpDateTimeS As String
    On Error Resume Next
    
    
    tmpDateTime0 = GetGFSShapeTime(mainShp)
'    tmpDateTime = DateAdd("n", Int(t), tmpDateTime)
'    tmpDateTime = DateAdd("n", frm_Command.GetCommandTime, tmpDateTime0)
    tmpDateTime = frm_Command.GetCommandDateTime
    
'    Debug.Print (CDate(tmpDateTime0))
'    Debug.Print (CDate(tmpDateTime))
                    
    If tmpDateTime > 0 Then
        setcellval shp, "Prop.Unit", cellVal(mainShp, "Prop.Unit", visUnitsString)
        setcellval shp, "Prop.Call", cellVal(mainShp, "Prop.Call", visUnitsString)
    
    
        If frm_Command.isOk Then
            '�������� ������ ����� ������������, ���� �� ����� ������� ��� ����� "��"
            If shp.CellExists("User.InPage", 0) = 0 Then
                shp.AddNamedRow visSectionUser, "InPage", visRowUser
                setcellval shp, "User.InPage", 1
                setcellval shp, "User.InPage.Prompt", "+"
            End If

            tmpDateTime0S = ToTimeString(tmpDateTime0)
            tmpDateTimeS = ToTimeString(tmpDateTime)
            SetCellFrml shp, "Prop.SetTime", tmpDateTimeS
            SetCellFrml shp, "Prop.FormingTime", tmpDateTime0S
            SetCellFrml shp, "Prop.ArrivalTime", tmpDateTimeS
        End If
    End If
    
End Function

Private Function GetShpData_0(ByRef mainShp As Visio.Shape, shp As Visio.Shape) As Variant
Dim tmpDateTime0 As Double
Dim tmpDateTime As Double
Dim tmpDateTime0S As String
Dim tmpDateTimeS As String
    On Error Resume Next
    
    
    tmpDateTime0 = GetGFSShapeTime(mainShp)
    tmpDateTime = tmpDateTime0

    
'    Debug.Print (CDate(tmpDateTime0))
'    Debug.Print (CDate(tmpDateTime))
                    
    If tmpDateTime > 0 Then
        setcellval shp, "Prop.Unit", cellVal(mainShp, "Prop.Unit", visUnitsString)
        setcellval shp, "Prop.Call", cellVal(mainShp, "Prop.Call", visUnitsString)
    
        '�������� ������ ����� ������������, ���� �� ����� ������� ��� ����� "��"
        If shp.CellExists("User.InPage", 0) = 0 Then
            shp.AddNamedRow visSectionUser, "InPage", visRowUser
            setcellval shp, "User.InPage", 1
            setcellval shp, "User.InPage.Prompt", "+"
        End If

        tmpDateTime0S = ToTimeString(tmpDateTime0)
        tmpDateTimeS = ToTimeString(tmpDateTime)
        SetCellFrml shp, "Prop.SetTime", tmpDateTimeS
        SetCellFrml shp, "Prop.FormingTime", tmpDateTime0S
        SetCellFrml shp, "Prop.ArrivalTime", tmpDateTimeS
    End If
    
End Function





'Private Function GetGFSShapeTime(ByRef shp As Visio.Shape) As Double
'
'    GetGFSShapeTime = cellval(shp, "Prop.SetTime", visDate)
'    If GetGFSShapeTime > 0 Then Exit Function
'
'    GetGFSShapeTime = cellval(shp, "Prop.FormingTime", visDate)
'    If GetGFSShapeTime > 0 Then Exit Function
'
'    GetGFSShapeTime = cellval(shp, "Prop.ArrivalTime", visDate)
'    If GetGFSShapeTime > 0 Then Exit Function
'
'GetGFSShapeTime = 0
'End Function

Private Function ToTimeString(ByVal dblVal As Double) As String
Dim tmp As String
    tmp = Replace(str(dblVal), ",", ".")
    ToTimeString = "DATETIME(" & tmp & ")"
End Function
