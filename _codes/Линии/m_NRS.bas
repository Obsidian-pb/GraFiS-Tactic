Attribute VB_Name = "m_NRS"
Option Explicit

'-------------------������ ��� �������� �������� � ������� �� ������� �������-�������� ������--------------------
Private shapesInNRS As Collection

'---���������� ������� ��������
'Const ccs_InIdent = "Connections.GFS_In"
'Const ccs_OutIdent = "Connections.GFS_Ou"
Const vb_ShapeType_Other = 0    '������
Const vb_ShapeType_Hose = 1     '������
Const vb_ShapeType_PTV = 2      '���
Const vb_ShapeType_Razv = 3     '������������
Const vb_ShapeType_Tech = 4     '�������
Const vb_ShapeType_VsasSet = 5  '����������� ����� � ������
Const vb_ShapeType_GE = 6       '�������������

Const CP_GrafisVersion = 1      '������ ������

'-----------------------����� �������� �������-----------------------
Public PA_Count As Integer
Public MP_Count As Integer
Public Ship_Count As Integer
Public Train_Count As Integer

Public Hose38_Count As Integer
Public Hose51_Count As Integer
Public Hose66_Count As Integer
Public Hose77_Count As Integer
Public Hose89_Count As Integer
Public Hose110_Count As Integer
Public Hose150_Count As Integer
Public Hose200_Count As Integer
Public Hose250_Count As Integer
Public Hose300_Count As Integer
Public OtherHoses_Count As Integer

Public NapHoses_Lenight As Integer
Public VsasHoses_Lenight As Integer

Public Hose77NV_Count As Integer
Public Hose125NV_Count As Integer
Public Hose150NV_Count As Integer
Public Hose200NV_Count As Integer

Public Razv_Count As Integer
Public VS_Count As Integer
Public Collector As Integer
Public GE_Count As Integer
Public PS_Count As Integer
Public VsasSetc_Count As Integer
Public Kol_Count As Integer
Public PV_Count As Integer

Public StvA_Count As Integer
Public StvB_Count As Integer
Public StvLaf_Count As Integer
Public StvPen_Count As Integer
Public StvGPS_Count As Integer

Public PodOut As Double
Public PodIn As Double
Public HosesValue As Double
Public WaterValue As Double '����� ���� � ��������

Public PG_Count As Integer
Public PW_Count As Integer
Public PK_Count As Integer
Public WaterContainers_Count As Integer
Public WaterContainers_Value As Double
Public WaterOpen_Value As Double

Private PW_Value As Double




Public Sub GESystemTest(ShpObj As Visio.Shape)
'�������� ��������� ��������� �������� � �������-�������� �������
    
    On Error GoTo EX
    
    '---�������� ������ ��� ��������� ����� � ���
    Set shapesInNRS = New Collection
    
    '---��������� ��������� �������� ��������� � ���
        GetTechShapeForGESystem ShpObj
    
    '---����������� ������ � ���������
        NRS_Analize
    '---��������� �����
        CreateReport
        
    Set shapesInNRS = Nothing
Exit Sub
EX:
    Set shapesInNRS = Nothing
End Sub

Private Sub GetTechShapeForGESystem(ByRef shp As Visio.Shape)
'��������� ��������� ����� ����������� � ���
Dim Con As Connect
Dim sideShp As Visio.Shape

    For Each Con In shp.Connects
        If Not IsShapeAllreadyChecked(Con.ToSheet) Then
            shapesInNRS.Add Con.ToSheet
            GetTechShapeForGESystem Con.ToSheet
        End If
    Next Con
    For Each Con In shp.FromConnects
        If Not IsShapeAllreadyChecked(Con.FromSheet) Then
            shapesInNRS.Add Con.FromSheet
            GetTechShapeForGESystem Con.FromSheet
        End If
    Next Con

End Sub

Private Function IsShapeAllreadyChecked(ByRef shp As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ ��� ������� � ����, ���� ���
Dim colShape As Visio.Shape

    For Each colShape In shapesInNRS
        If colShape = shp Then
            IsShapeAllreadyChecked = True
            Exit Function
        End If
    Next colShape
    
IsShapeAllreadyChecked = False
End Function

Private Sub CreateReport()
'����� ��������� � ������� ����� �����
Dim totalStr As String
Dim FlowOut As Double '�������� �������� ��������
Dim DischargeTime As Double
    
    If PodOut > 0 Then totalStr = totalStr & "����� ������ ������� - " & PodOut & " �/�" & Chr(10)
    If PodIn > 0 Then totalStr = totalStr & "����� ����� ���� - " & PodIn & " �/�" & Chr(10)
    If HosesValue > 0 Then totalStr = totalStr & "����� ���� � ������� - " & HosesValue & " �" & Chr(10)
    If WaterValue > 0 Then totalStr = totalStr & "����� ���� � �������� ��� - " & WaterValue & " �" & Chr(10)
    If PW_Value > 0 Then totalStr = totalStr & "����� ��������� ���� � �������� �������� - " & PW_Value & " �" & Chr(10)
    If WaterOpen_Value > 0 Then totalStr = totalStr & "����� ��������� ���� � �������� �������������� - " & WaterOpen_Value & " �" & Chr(10)

    If PodOut > PodIn Then
        FlowOut = PodOut - PodIn
        DischargeTime = ((WaterValue + PW_Value + WaterOpen_Value - HosesValue) / FlowOut) / 60
        If Int(DischargeTime) > 0 Then
            totalStr = totalStr & "��������� ����� ������ ������� - " & _
                     Int(DischargeTime) & ":" & Int((DischargeTime - Int(DischargeTime)) * 60) _
                     & Chr(10)
        Else
            totalStr = totalStr & "������ ������� ������� ������ �������!" & Chr(10)
        End If
    Else
        totalStr = totalStr & "��������� ����� ������ ������� - ����������" & Chr(10)
    End If
    
    If PA_Count > 0 Then totalStr = totalStr & "�������� ����������� - " & PA_Count & Chr(10)
    If MP_Count > 0 Then totalStr = totalStr & "�������� �������� - " & MP_Count & Chr(10)
    If Train_Count > 0 Then totalStr = totalStr & "�������� ������� - " & Train_Count & Chr(10)
    If Ship_Count > 0 Then totalStr = totalStr & "�������� ����� - " & Ship_Count & Chr(10)
    
    If Hose38_Count > 0 Then totalStr = totalStr & "�������� ������ 38 �� - " & Hose38_Count & Chr(10)
    If Hose51_Count > 0 Then totalStr = totalStr & "�������� ������ 51 �� - " & Hose51_Count & Chr(10)
    If Hose66_Count > 0 Then totalStr = totalStr & "�������� ������ 66 �� - " & Hose66_Count & Chr(10)
    If Hose77_Count > 0 Then totalStr = totalStr & "�������� ������ 77 �� - " & Hose77_Count & Chr(10)
    If Hose89_Count > 0 Then totalStr = totalStr & "�������� ������ 89 �� - " & Hose89_Count & Chr(10)
    If Hose110_Count > 0 Then totalStr = totalStr & "�������� ������ 110 �� - " & Hose110_Count & Chr(10)
    If Hose150_Count > 0 Then totalStr = totalStr & "�������� ������ 150 �� - " & Hose150_Count & Chr(10)
    If Hose200_Count > 0 Then totalStr = totalStr & "�������� ������ 200 �� - " & Hose200_Count & Chr(10)
    If Hose250_Count > 0 Then totalStr = totalStr & "�������� ������ 250 �� - " & Hose250_Count & Chr(10)
    If Hose300_Count > 0 Then totalStr = totalStr & "�������� ������ 300 �� - " & Hose300_Count & Chr(10)
    If OtherHoses_Count > 0 Then totalStr = totalStr & "������ �������� ������ - " & OtherHoses_Count & Chr(10)
    If NapHoses_Lenight > 0 Then totalStr = totalStr & "����� �������� �������� ����� - " & NapHoses_Lenight & " � " & Chr(10)
    If Hose77NV_Count > 0 Then totalStr = totalStr & "�������-����������� ������ 77 �� - " & Hose77NV_Count & Chr(10)
    If Hose125NV_Count > 0 Then totalStr = totalStr & "����������� ������ 125 �� - " & Hose125NV_Count & Chr(10)
    If Hose150NV_Count > 0 Then totalStr = totalStr & "����������� ������ 150 �� - " & Hose150NV_Count & Chr(10)
    If Hose200NV_Count > 0 Then totalStr = totalStr & "����������� ������ 200 �� - " & Hose200NV_Count & Chr(10)
    If VsasHoses_Lenight > 0 Then totalStr = totalStr & _
            "����� ����������� (�������-�����������) �������� ����� - " & VsasHoses_Lenight & " � " & Chr(10)
    
    If StvB_Count > 0 Then totalStr = totalStr & "������� � - " & StvB_Count & Chr(10)
    If StvA_Count > 0 Then totalStr = totalStr & "������� � - " & StvA_Count & Chr(10)
    If StvLaf_Count > 0 Then totalStr = totalStr & "�������� ������� - " & StvLaf_Count & Chr(10)
    If StvPen_Count > 0 Then totalStr = totalStr & "������ ������� - " & StvPen_Count & Chr(10)
'    If StvGPS_Count > 0 Then totalStr = totalStr & "������ ������� ������� - " & StvGPS_Count & Chr(10)
    
    If Razv_Count > 0 Then totalStr = totalStr & "������������ - " & Razv_Count & Chr(10)
    If VS_Count > 0 Then totalStr = totalStr & "������������� - " & VS_Count & Chr(10)
    If Collector > 0 Then totalStr = totalStr & "����������� - " & VS_Count & Chr(10)
    If GE_Count > 0 Then totalStr = totalStr & "��������������� - " & GE_Count & Chr(10)
    If PS_Count > 0 Then totalStr = totalStr & "�������������� - " & PS_Count & Chr(10)
    If PV_Count > 0 Then totalStr = totalStr & "������ ������� - " & PV_Count & Chr(10)
    If VsasSetc_Count > 0 Then totalStr = totalStr & "����������� ����� - " & VsasSetc_Count & Chr(10)
    If Kol_Count > 0 Then totalStr = totalStr & "������� - " & Kol_Count & Chr(10)
    
    If PG_Count > 0 Then totalStr = totalStr & "������������ �������� ��������� - " & PG_Count & Chr(10)
    If PW_Count > 0 Then totalStr = totalStr & "������������ �������� �������� - " & PW_Count & Chr(10)
    If PK_Count > 0 Then totalStr = totalStr & "������������ �������� ������ - " & PK_Count & Chr(10)
    If WaterContainers_Count > 0 Then totalStr = totalStr & "������������� �������� ��� ���� - " & WaterContainers_Count & Chr(10)
    
    If PW_Count > 0 Then totalStr = totalStr & "�������� �������� - " & PW_Count & Chr(10)

    
    MsgBox totalStr, vbOKOnly, "������ �������-�������� �������"
End Sub



Public Sub Test()
    GESystemTest Application.ActiveWindow.Selection(1)
End Sub

Public Sub ClearVaraibles()
'������� ��� ����������
     PA_Count = 0
     MP_Count = 0
     Train_Count = 0
     Ship_Count = 0
    
     Hose38_Count = 0
     Hose51_Count = 0
     Hose66_Count = 0
     Hose77_Count = 0
     Hose89_Count = 0
     Hose110_Count = 0
     Hose150_Count = 0
     Hose200_Count = 0
     Hose250_Count = 0
     Hose300_Count = 0
     OtherHoses_Count = 0
    
     NapHoses_Lenight = 0
     VsasHoses_Lenight = 0
    
     Hose77NV_Count = 0
     Hose125NV_Count = 0
     Hose150NV_Count = 0
     Hose200NV_Count = 0
    
     Razv_Count = 0
     VS_Count = 0
     Collector = 0
     GE_Count = 0
     PS_Count = 0
     VsasSetc_Count = 0
     Kol_Count = 0
     PV_Count = 0
    
     StvA_Count = 0
     StvB_Count = 0
     StvLaf_Count = 0
     StvPen_Count = 0
     StvGPS_Count = 0
     
     PodOut = 0
     PodIn = 0
     HosesValue = 0
     WaterValue = 0
'     WaterContainers_Value = 0
    
     PG_Count = 0
     PW_Count = 0
     PK_Count = 0
     WaterContainers_Count = 0
     
     PW_Count = 0
     PW_Value = 0
     WaterOpen_Value = 0
End Sub

'----------------------------------------��������� �������-----------------------------------------------
Public Sub NRS_Analize()
'��������� ������� �������-�������� �������
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer

    On Error GoTo EX

'---������� ��� ����������
    ClearVaraibles

'---���������� ��� ������ � � ������ ���� ������ �������� ������� ������ ����������� �
    For Each vsO_Shape In shapesInNRS
        If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then '�������� �� ������ ������� ������
            If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then  '��������� ������ ������
                vsi_ShapeIndex = vsO_Shape.Cells("User.IndexPers")   '���������� ������ ������ ������
                
                '---��������� �������������� (���� ������ �����������  �������)
                If IsNotManeuwer(vsO_Shape) Then
                    
                
                    '---����� �������� (����������� ��� ���������� ����� �����, �� ������ ������)
                    '---������ �� ��������� ������
                    If vsO_Shape.CellExists("User.GFS_OutLafet", 0) = True Then
                        PodOut = PodOut + vsO_Shape.Cells("User.GFS_OutLafet").Result(visNumber)
                    End If

                
                    Select Case vsi_ShapeIndex
                    '---�������� ����������-----------
                        Case Is = indexPers.ipAC '������������
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipANR '���
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipPNS '���
                            PA_Count = PA_Count + 1
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAA 'AA
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAV '��
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAKT '���
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAGVT '����
                            PA_Count = PA_Count + 1
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 2 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAR '��
                            PA_Count = PA_Count + 1
                        Case Is = indexPers.ipACL '���
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipACKP '����
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                        Case Is = indexPers.ipAPP '���
                            PA_Count = PA_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                            If vsO_Shape.Cells("Actions.WaterCollect.Checked") Then VS_Count = VS_Count + 1 '����������� ����������� �� �/�
                    
                    '---������ �������� �������----------
                        Case Is = indexPers.ipPoezd '������
                            Train_Count = Train_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.Water").Result(visNumber)
                        Case Is = indexPers.ipMotoPump '���������
                            MP_Count = MP_Count + 1
                        Case Is = indexPers.ipKorabl '�������
                            Ship_Count = Ship_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                        Case Is = indexPers.ipKater '�����
                            Ship_Count = Ship_Count + 1
                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                        Case Is = indexPers.ipMashNaGusenicah '������ �� ���������� ����
                        Case Is = indexPers.ipTank '�����
            
                    '---���-----------------------------------
                        Case Is = indexPers.ipStvolRuch '������ �������
                            If vsO_Shape.Cells("User.DiameterIn").Result(visNumber) <= 50 Then
                                StvB_Count = StvB_Count + 1
                            Else
                                StvA_Count = StvA_Count + 1
                            End If
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolRuchPena '������ ������
                            StvPen_Count = StvPen_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafVoda '�������� �������
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafPena '�������� ������
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipStvolLafVoda '�������� ������� �������
                            StvLaf_Count = StvLaf_Count + 1
                            PodOut = PodOut + vsO_Shape.Cells("User.PodOut").Result(visNumber)
                        Case Is = indexPers.ipRazvetvlenie '������������
                            Razv_Count = Razv_Count + 1
                        Case Is = indexPers.ipVodosbornik 'Vodosbornik
                            VS_Count = VS_Count + 1
                        Case Is = indexPers.ipKollector 'Collector
                            Collector = Collector + 1
                        Case Is = indexPers.ipPenopodemnik '�������������
                        Case Is = indexPers.ipGrebenkaPGenerat '��������
                        Case Is = indexPers.ipPennayaVstavka '������ �������
                            PV_Count = PV_Count + 1
    
            
                    '---�������������
                        Case Is = indexPers.ipPG '��
                            PG_Count = PG_Count + 1
                        Case Is = indexPers.ipPW '��
                            PW_Count = PW_Count + 1
                        Case Is = indexPers.ipPK '��
                            PK_Count = PK_Count + 1
                        Case Is = indexPers.ipOtkritiyVodoistochnik '������
                        Case Is = indexPers.ipPirs '����
                        Case Is = indexPers.ipBashna '�����
                            
                    '---����� ����
                        Case Is = indexPers.ipHydroelevator '�������������
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "��" Then
                                GE_Count = GE_Count + 1
                                GetWaterSourceData vsO_Shape
                            End If
                        Case Is = indexPers.ipPenosmesitelPerenosn '�������������
                            PS_Count = PS_Count + 1
                        Case Is = indexPers.ipVsasLineWithSetk  '����������� ����� � ������
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "��" Then
                                VsasSetc_Count = VsasSetc_Count + 1
    '                            PodIn = PodIn + vsO_Shape.Cells("Prop.PodIn").Result(visNumber)
                                GetWaterSourceData vsO_Shape
                            End If
                        Case Is = indexPers.ipKolonka '�������
                            If CellVal(vsO_Shape, "Prop.GetingWater", visUnitsString) = "��" Then
                                Kol_Count = Kol_Count + 1
                                PodIn = PodIn + vsO_Shape.Cells("Prop.FlowCurrent").Result(visNumber)
                            End If
                        Case Is = indexPers.ipEmkost '������� ��� ����
                            WaterContainers_Count = WaterContainers_Count + 1
                            WaterValue = WaterValue + vsO_Shape.Cells("Prop.WaterContainerValue").Result(visNumber) * _
                                    vsO_Shape.Cells("Prop.OstKoeff").Result(visNumber)
                       
                    '---�����
                        Case Is = indexPers.ipRukavLineNapor '�������� �����
                            If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "���" _
                                And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "�����(4�)" Then
                                Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
                                    Case Is = 38
                                        Hose38_Count = Hose38_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 51
                                        Hose51_Count = Hose51_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 66
                                        Hose66_Count = Hose66_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 77
                                        Hose77_Count = Hose77_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 89
                                        Hose89_Count = Hose89_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 110
                                        Hose110_Count = Hose110_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 150
                                        Hose150_Count = Hose150_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 200
                                        Hose200_Count = Hose200_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 250
                                        Hose250_Count = Hose250_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 300
                                        Hose300_Count = Hose300_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                        
                                End Select
                                '---����� �����
                                If vsO_Shape.CellExists("User.TotalLenight", 0) = True Then '����� �� ������ ������ "User.TotalLenight"
                                    NapHoses_Lenight = NapHoses_Lenight + vsO_Shape.Cells("User.TotalLenight")
                                Else
                                    NapHoses_Lenight = NapHoses_Lenight + vsO_Shape.Cells("Prop.LineLenightHose")
                                End If
                                '---����� ���� � ������
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
                            If vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "�����(4�)" Then
                                OtherHoses_Count = OtherHoses_Count + 1
                                '---����� ���� � ������
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
                        Case Is = indexPers.ipRukavLineVsas '����������� ����� ��� �������-�����������
                            If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "���" _
                                And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "�����(4�)" Then
                                Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
                                    Case Is = 77
                                        Hose77NV_Count = Hose77NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 125
                                        Hose125NV_Count = Hose125NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 150
                                        Hose150NV_Count = Hose150NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                    Case Is = 200
                                        Hose200NV_Count = Hose200NV_Count + vsO_Shape.Cells("User.HosesNeed").Result(visNumber)
                                End Select
                                '---����� �����
                                If vsO_Shape.CellExists("User.TotalLenight", 0) = True Then '����� �� ������ ������ "User.TotalLenight"
                                    VsasHoses_Lenight = VsasHoses_Lenight + vsO_Shape.Cells("User.TotalLenight")
                                Else
                                    VsasHoses_Lenight = VsasHoses_Lenight + vsO_Shape.Cells("Prop.LineLenightHose")
                                End If
                                '---����� ���� � ������
                                HosesValue = HosesValue + vsO_Shape.Cells("Prop.LineValue").Result(visNumber)
                            End If
    
                    End Select
                End If
            End If
        End If
    Next vsO_Shape



Exit Sub
EX:
    SaveLog Err, "NRS_Analize"
    
End Sub

Private Function IsNotManeuwer(ByRef shp As Visio.Shape) As Boolean
'������� ���������� ������, ���� ������ �� �����������, ��� ����� �������� � ��� ������ �����������
    If shp.CellExists("Actions.MainManeure.Checked", 0) = True Then
        If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
            IsNotManeuwer = True
        Else
            IsNotManeuwer = False
        End If
        Exit Function
    End If
IsNotManeuwer = True
End Function

Private Sub GetWaterSourceData(ByRef shp As Visio.Shape)
'�������� �������� � ��������������� ������� �� ����� ����������� ����� � ������ ��� ��������������
Dim shpWS As Visio.Shape
Dim shpWSIndex As Long
Dim tmpVal As Long
    
    shpWSIndex = CellVal(shp, "User.WSShapeID")
    If shpWSIndex > 0 Then
        Set shpWS = Application.ActivePage.Shapes.ItemFromID(shpWSIndex)
        If IsGFSShapeWithIP(shpWS, indexPers.ipPW) Then
            PW_Count = PW_Count + 1
            
            tmpVal = CellVal(shpWS, "Prop.WaterValue") * 1000
            PW_Value = PW_Value + tmpVal
        End If
        If IsGFSShapeWithIP(shpWS, indexPers.ipOtkritiyVodoistochnik) Then
            If CellVal(shpWS, "Prop.Type", visUnitsString) = "�������������� �����" Then
                PodIn = PodIn + CellVal(shp, "Prop.PodIn")
                PodIn = PodIn + CellVal(shp, "Prop.PodFromOuter")
            Else
                tmpVal = CellVal(shpWS, "Prop.Value") * 1000 * 0.9  '� ������ ������������ �������������
                WaterOpen_Value = WaterOpen_Value + tmpVal
            End If
        End If
    End If
    
End Sub












