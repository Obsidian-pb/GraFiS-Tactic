Attribute VB_Name = "m_WorkWithConnections"
Option Explicit
'--------------------------------------------------------------������ �������� ��������� ���������� ��� � �������� �������--------------------------------------
Private cpO_InShape As Visio.Shape, cpO_OutShape As Visio.Shape '������ � ������� ������ ����� � �������, ��������������


'---���������� �������
Const ccs_InIdent = "Connections.GFS_In"
Const ccs_OutIdent = "Connections.GFS_Ou"
Const vb_ShapeType_Other = 0                '������
Const vb_ShapeType_Hose = 1                 '������
Const vb_ShapeType_PTV = 2                  '���
Const vb_ShapeType_Razv = 3                 '������������
Const vb_ShapeType_Tech = 4                 '�������
Const vb_ShapeType_VsasSet = 5              '����������� ����� � ������
Const vb_ShapeType_GE = 6                   '�������������
Const vb_ShapeType_WaterContainer = 7       '������� �������



'----------------------------------------��������� ������ � ������������-----------------------------------------------
Public Sub ConnectionsRefresh(ShpObj As Visio.Shape)
'��������� �������� ����������� ������� � ������� ������
Dim Conn As Visio.Connect
Dim i As Integer

On Error Resume Next  ' �� ������ - ������ ������������ - ��� ����, ����� �� ��������� ��������� �� ������ �������!!!

'---������� �������� ��� ������������� ��������� ������
'    If ShpObj.CellExists("User.GFS_OutLafet", 0) = True Then
'        ShpObj.Cells("User.GFS_OutLafet").FormulaU = 0
'        ShpObj.Cells("User.GFS_OutLafet.Prompt").FormulaU = 0
'    End If
    
'---���������� i
    i = 0

'---������� �������� � ������� ������ Scratch - ���������� ���������� � ���������� ������������
'    Do While ShpObj.CellsSRCExists(visSectionScratch, i, 0, 0) = True  '��� ������ Scratch!!!
    Do While ShpObj.CellsSRCExists(visSectionConnectionPts, i, 0, 0) = True  '��� ������ Connections!!!
        If Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_Ou" Then
'                ShpObj.CellsSRC(visSectionScratch, i, 4).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
                ShpObj.CellsSRC(visSectionScratch, i, 3).FormulaU = 0
        ElseIf Left(ShpObj.Section(visSectionConnectionPts).row(i).Name, 6) = "GFS_In" Then
                ShpObj.CellsSRC(visSectionScratch, i, 2).FormulaU = 0
        End If
        ShpObj.CellsSRC(visSectionScratch, i, 0).FormulaU = 0
        i = i + 1
        If i > 100 Then Exit Do
    Loop

'---��������� ���������� �������
    For Each Conn In ShpObj.FromConnects
        Ps_ConnectionAdd Conn
    Next Conn

Set Conn = Nothing
End Sub


Public Sub Ps_ConnectionAdd(ByRef aO_Conn As Visio.Connect)
'��������� ��������������� ������������ ���������� �����
Dim vO_FromShape As Visio.Shape, vO_ToShape As Visio.Shape
Dim vi_InRowNumber As Integer, vi_OutRowNumber As Integer

On Error GoTo EndSub
    
'---���������� ���� ������ ���� ���������
    Set vO_FromShape = aO_Conn.FromSheet
    Set vO_ToShape = aO_Conn.ToSheet

'---���������, �������� �� ����������� ������ �������� ������
    If vO_FromShape.CellExists("User.IndexPers", 0) = False Or _
        vO_ToShape.CellExists("User.IndexPers", 0) = False Then Exit Sub '---��������� �������� �� ������ _
                                                                                �������� �����
'---���������, �������� �� ����������� ������ ���������� ���
    If f_IdentShape(vO_FromShape.Cells("User.IndexPers").Result(visNumber)) = 0 Or _
        f_IdentShape(vO_ToShape.Cells("User.IndexPers").Result(visNumber)) = 0 Then Exit Sub

'---�������������� �������� � ����������� ������ - ��� ���������� ������� � ���!!!
    '---��� From ������
    If Left(aO_Conn.FromCell.Name, 18) = ccs_InIdent Then
        Set cpO_InShape = aO_Conn.FromSheet
        Set cpO_OutShape = aO_Conn.ToSheet
        vi_InRowNumber = aO_Conn.FromCell.row
        vi_OutRowNumber = aO_Conn.ToCell.row
    ElseIf Left(aO_Conn.FromCell.Name, 18) = ccs_OutIdent Then
        Set cpO_InShape = aO_Conn.ToSheet
        Set cpO_OutShape = aO_Conn.FromSheet
        vi_InRowNumber = aO_Conn.ToCell.row
        vi_OutRowNumber = aO_Conn.FromCell.row
    End If
    '---��� �� ������
    If Left(aO_Conn.ToCell.Name, 18) = ccs_InIdent Then
        Set cpO_InShape = aO_Conn.ToSheet
        Set cpO_OutShape = aO_Conn.FromSheet
        vi_InRowNumber = aO_Conn.ToCell.row
        vi_OutRowNumber = aO_Conn.FromCell.row
    ElseIf Left(aO_Conn.ToCell.Name, 18) = ccs_OutIdent Then
        Set cpO_InShape = aO_Conn.FromSheet
        Set cpO_OutShape = aO_Conn.ToSheet
        vi_InRowNumber = aO_Conn.FromCell.row
        vi_OutRowNumber = aO_Conn.ToCell.row
    End If
    '---� ������, ���� ��� ������ - ������
    If vO_FromShape.Cells("User.IndexPers") = 100 And _
        vO_ToShape.Cells("User.IndexPers") = 100 Then
        '---��������� � ����� ������ ������ �������� ����� ������
        If aO_Conn.ToSheet.Cells("Scratch.D1") > aO_Conn.FromSheet.Cells("Scratch.D1") Then
            Set cpO_InShape = aO_Conn.ToSheet
            Set cpO_OutShape = aO_Conn.FromSheet
            vi_InRowNumber = aO_Conn.ToCell.row
            vi_OutRowNumber = aO_Conn.FromCell.row
        Else
            '---��������� � ����� ������ ������ ID ���� (�������� �����)
            If aO_Conn.ToSheet.ID > aO_Conn.FromSheet.ID Then
                Set cpO_InShape = aO_Conn.ToSheet
                Set cpO_OutShape = aO_Conn.FromSheet
                vi_InRowNumber = aO_Conn.ToCell.row
                vi_OutRowNumber = aO_Conn.FromCell.row
            Else
                Set cpO_InShape = aO_Conn.FromSheet
                Set cpO_OutShape = aO_Conn.ToSheet
                vi_InRowNumber = aO_Conn.FromCell.row
                vi_OutRowNumber = aO_Conn.ToCell.row
            End If
        End If
    End If

    '---��������� ��������� ���������� ������ � �������
       ps_LinkShapes vi_InRowNumber, vi_OutRowNumber

Exit Sub

EndSub:
    Resume Next
'    Debug.Print Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "Ps_ConnectionAdd"
    Set cpO_InShape = Nothing
    Set cpO_OutShape = Nothing
End Sub

Private Sub ps_LinkShapes(ByVal ai_InRowNumber As Integer, ByVal ai_OutRowNumber As Integer)
'���������� ��������� - ��������� ������ � ����������� �������
Dim vi_IPInShape, vi_IPOutShape As Integer
Dim vb_InShapeType, vb_OutShapeType As Byte
Dim i As Integer
Dim vs_Formula As String

On Error GoTo EX

'---��������� ��� �������� ����������� ������
    '---�������������IndexPers ��� ������ �� �����
    vi_IPInShape = cpO_InShape.Cells("User.IndexPers")
    vi_IPOutShape = cpO_OutShape.Cells("User.IndexPers")
    '---���������
        '---��� ����������� ������
        vb_InShapeType = f_IdentShape(vi_IPInShape)
        '---��� �������� ������
        vb_OutShapeType = f_IdentShape(vi_IPOutShape)
        
'---� ����������� �� ���� ����������� ����� ������� ��������� ����������
    '---�����->���
        If vb_OutShapeType = vb_ShapeType_Hose And (vb_InShapeType = vb_ShapeType_PTV Or vb_InShapeType = vb_ShapeType_Razv Or vb_InShapeType = vb_ShapeType_GE) Then
            '����� �������� �� ��������� ���������
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula      '������ �������� ���
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---���������� ��� �������� ������� ������� � ������� ������:
            If vi_IPInShape = 36 Or vi_IPInShape = 37 Or vi_IPInShape = 39 Or vi_IPInShape = 35 Then
            '---���������, ��� � ������ ������������ ������ �� ����
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            End If
            '---����������� ������������� ����� � ������������� ������
            If cpO_InShape.CellExists("Actions.MainManeure", 0) = True Then
                cpO_OutShape.Cells("Prop.ManeverHose").Formula = "INDEX(Sheet." & cpO_InShape.ID _
                    & "!Actions.MainManeure.Checked" & ";Prop.ManeverHose.Format)"
            End If
            '---����������� ����� ��� � ������ �������� �����
            cpO_InShape.Cells("LineColor").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
            cpO_InShape.Cells("FillForegnd").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
            cpO_InShape.Cells("Char.Color").Formula = "Sheet." & cpO_OutShape.ID & "!LineColor"
        End If
    '---���->�����
        If (vb_OutShapeType = vb_ShapeType_PTV Or vb_OutShapeType = vb_ShapeType_Razv Or vb_OutShapeType = vb_ShapeType_GE) And vb_InShapeType = vb_ShapeType_Hose Then
            
            '��������� �������� ������ ������
            cpO_InShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C" & ai_OutRowNumber + 1
            
            '���������� ������ ������ ������������
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
            '��������� ��������� ������� �� ������
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, visScratchB).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
        End If
    '---�����->�����
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Hose Then
            '����
'            cpO_OutShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            '�����
            cpO_InShape.Cells("Scratch.A1").Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            cpO_OutShape.Cells("Scratch.B1").Formula = "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.LineTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            cpO_OutShape.Cells("Prop.ManeverHose").FormulaU = "Sheet." & cpO_InShape.ID & "!Prop.ManeverHose"
        End If
    '---�����->������� ��������
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_Tech Then
            '����
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
'            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            '�����
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).Formula = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.ArrivalTime"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---���������, ��� � ������ ������������ ������ �� ����
            If cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            ElseIf cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "2m"
            End If
            '---���������, �� �������� �� ����� ���� �� ���� �� ���, �������� ������
            If SelfWaterGetCheck(cpO_OutShape) Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = 0
            End If
        End If
    '---������� ��������->�����
        If vb_OutShapeType = vb_ShapeType_Tech And vb_InShapeType = vb_ShapeType_Hose Then
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 3).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            cpO_InShape.Cells("Scratch.A1").Formula = _
                "Sheet." & cpO_OutShape.ID & "!Scratch.C" & ai_OutRowNumber + 1
                
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            '---���������, ��� � ������ ������������ ������ �� �����
            If cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "1m"
            ElseIf cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "2m"
            End If
        End If
    '---����������� �����->������� ��������
        If vb_OutShapeType = vb_ShapeType_VsasSet And vb_InShapeType = vb_ShapeType_Tech Then
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
'            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
            cpO_OutShape.Cells("Prop.Unit").Formula = "Sheet." & cpO_InShape.ID & "!Prop.Unit"
            '---���������, ��� � ������ ������������ ������ �� ����
            If cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
            ElseIf cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "2m"
            End If
        End If
    '---�����->�������
        If vb_OutShapeType = vb_ShapeType_Hose And vb_InShapeType = vb_ShapeType_WaterContainer Then
'            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.C" _
'                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.C" & ai_InRowNumber + 1 & ")"
'            cpO_OutShape.Cells("Scratch.A1").FormulaU = vs_Formula
            
            cpO_InShape.Cells("Scratch.A" & ai_InRowNumber + 1).FormulaU = "Sheet." & cpO_OutShape.ID & "!Scratch.C1"
            vs_Formula = "IF(ISERR(Sheet." & cpO_InShape.ID & "!Scratch.D" _
                & ai_InRowNumber + 1 & "),0," & "Sheet." & cpO_InShape.ID & "!Scratch.D" & ai_InRowNumber + 1 & ")"
            cpO_OutShape.Cells("Scratch.B1").FormulaU = vs_Formula
            cpO_OutShape.Cells("User.FlowToShape").Formula = """" & CStr(cpO_InShape.NameU) & """"
'            cpO_OutShape.Cells("Prop.LineTime").Formula = "Sheet." & cpO_InShape.ID & "!Prop.SetTime"
            cpO_InShape.Cells("Scratch.X" & CStr(ai_InRowNumber + 1)).Formula = "1m"
        End If
    '---�������->�����
        If vb_OutShapeType = vb_ShapeType_WaterContainer And vb_InShapeType = vb_ShapeType_Hose Then
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 4).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.C1"
'            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 5).Formula = _
'                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
            
            cpO_OutShape.CellsSRC(visSectionScratch, ai_OutRowNumber, 3).Formula = _
                "Sheet." & cpO_InShape.ID & "!Scratch.D1"
'            cpO_InShape.Cells("Scratch.A1").Formula =  ���������������� �����?
            
            cpO_InShape.Cells("User.FlowFromShape").Formula = """" & CStr(cpO_OutShape.NameU) & """"
            '---���������, ��� � ������ ������������ ������ �� �����
            If cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 0 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "1m"
            ElseIf cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Result(visNumber) = 1 Then
                cpO_OutShape.Cells("Scratch.X" & CStr(ai_OutRowNumber + 1)).Formula = "2m"
            End If
        End If
    Exit Sub
    
    
    
    
Exit Sub
EX:
    Debug.Print Err.description
'    Debug.Print vs_Formula
    SaveLog Err, "ps_LinkShapes"

End Sub


'----------------------------------------��������� ���������-----------------------------------------------
Private Function SelfWaterGetCheck(ByRef hoseShp As Visio.Shape) As Boolean
'����� ���������, �� �������������� �� ��������� � ���� ��������������, ����������� � ����������� ��������� ��� ����������
Dim con As Visio.Connect
Dim shpName1 As String
Dim shpName2 As String
    
'---��������� ����� ������ ���������� � ������ ������
    For Each con In hoseShp.Connects
        If con.FromCell.Name = "BeginX" Then
            shpName1 = con.ToSheet.NameU
        End If
        If con.FromCell.Name = "EndX" Then
            shpName2 = con.ToSheet.NameU
        End If
    Next con
    
'---���� ���������� ���� � �� �� ������ - ���������� ������ �����, ����
    If shpName1 = shpName2 Then
        SelfWaterGetCheck = True
    Else
        SelfWaterGetCheck = False
    End If
End Function


'----------------------------------------��������� �������-----------------------------------------------
Private Function f_IdentShape(ByVal ai_ShapeIP As Integer) As Integer
'������� �������������� ������ � ���������� �������� � ����
Dim Arr_PTVs(26, 1) As Integer
Dim i As Integer

'---��������� �������� IndexPers � ��������������� �� �����������
    Arr_PTVs(0, 0) = 34  '������� ������ �����
        Arr_PTVs(0, 1) = vb_ShapeType_PTV
    Arr_PTVs(1, 0) = 35  '������ �����
        Arr_PTVs(1, 1) = vb_ShapeType_PTV
    Arr_PTVs(2, 0) = 36  '�������� �������
        Arr_PTVs(2, 1) = vb_ShapeType_PTV
    Arr_PTVs(3, 0) = 37  '�������� ������
        Arr_PTVs(3, 1) = vb_ShapeType_PTV
    Arr_PTVs(4, 0) = 39  '������� �������� �����
        Arr_PTVs(4, 1) = vb_ShapeType_PTV
    Arr_PTVs(5, 0) = 40  '�������������
        Arr_PTVs(5, 1) = vb_ShapeType_GE
    Arr_PTVs(6, 0) = 42  '������������
        Arr_PTVs(6, 1) = vb_ShapeType_Razv
    Arr_PTVs(7, 0) = 45  '�������������
        Arr_PTVs(7, 1) = vb_ShapeType_PTV
    Arr_PTVs(8, 0) = 72  '�������
        Arr_PTVs(8, 1) = vb_ShapeType_PTV
    Arr_PTVs(9, 0) = 88  '����������� ����� � ������
        Arr_PTVs(9, 1) = vb_ShapeType_VsasSet
    Arr_PTVs(10, 0) = 100 '�������� �����
        Arr_PTVs(10, 1) = vb_ShapeType_Hose
    Arr_PTVs(11, 0) = 101 '����������� �����
        Arr_PTVs(11, 1) = vb_ShapeType_Hose
    Arr_PTVs(12, 0) = 1 '������������ ��������
        Arr_PTVs(12, 1) = vb_ShapeType_Tech
    Arr_PTVs(13, 0) = 2 '���
        Arr_PTVs(13, 1) = vb_ShapeType_Tech
    Arr_PTVs(14, 0) = 20 '�������� ����������
        Arr_PTVs(14, 1) = vb_ShapeType_Tech
    Arr_PTVs(15, 0) = 161 '���
        Arr_PTVs(15, 1) = vb_ShapeType_Tech
    Arr_PTVs(16, 0) = 162 '����
        Arr_PTVs(16, 1) = vb_ShapeType_Tech
    Arr_PTVs(17, 0) = 163 '���
        Arr_PTVs(17, 1) = vb_ShapeType_Tech
    Arr_PTVs(18, 0) = 8 '���
        Arr_PTVs(18, 1) = vb_ShapeType_Tech
    Arr_PTVs(19, 0) = 9 '��
        Arr_PTVs(19, 1) = vb_ShapeType_Tech
    Arr_PTVs(20, 0) = 20 '��
        Arr_PTVs(20, 1) = vb_ShapeType_Tech
    Arr_PTVs(21, 0) = 13 '����
        Arr_PTVs(21, 1) = vb_ShapeType_Tech
    Arr_PTVs(22, 0) = 28 '���������
        Arr_PTVs(22, 1) = vb_ShapeType_Tech
    Arr_PTVs(23, 0) = 190 '������� � �����
        Arr_PTVs(23, 1) = vb_ShapeType_WaterContainer
    Arr_PTVs(24, 0) = 191 '������ �������
        Arr_PTVs(24, 1) = vb_ShapeType_Tech  '��� �� �������, �� ��� �����
    Arr_PTVs(25, 0) = 10 '���������� ������� �������
        Arr_PTVs(25, 1) = vb_ShapeType_Tech
    Arr_PTVs(26, 0) = 41 '�������������
        Arr_PTVs(26, 1) = vb_ShapeType_Tech  '��� �� �������, �� ��� �����
        
        
        
'---��������� �������� �� ���������
    f_IdentShape = vb_ShapeType_Other

'---��������� �������� �� ������
        For i = 0 To 26
            If ai_ShapeIP = Arr_PTVs(i, 0) Then f_IdentShape = Arr_PTVs(i, 1): Exit Function
        Next i

End Function


