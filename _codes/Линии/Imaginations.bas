Attribute VB_Name = "Imaginations"

Sub LenightSet(ShpObj As Visio.Shape, PRO As Integer)
'������� ��������� ���������� ���������� ���� ���������� ������ �������� ����� ������
    Dim LenightCalc As Integer
    
    On Error Resume Next
    
'    LenightCalc = ActivePage.Shapes.ItemFromID(PRO).LengthIU / 39.37 '��������� �� ������ � �����
'    Application.ActiveWindow.Page.Shapes.ItemFromID(PRO).Cells("User.LineLenight").FormulaForceU = LenightCalc
    LenightCalc = ShpObj.LengthIU / 39.37 '��������� �� ������ � �����
    ShpObj.Cells("User.LineLenight").FormulaForceU = LenightCalc
    
End Sub

Private Sub LenightSetInner(PRO As String)
'���������� ��������� ���������� ���������� ���� ���������� ������ �������� ����� ������
    Dim LenightCalc As Integer
    
    On Error Resume Next
    
    LenightCalc = ActivePage.Shapes(PRO).LengthIU / 39.37 '��������� �� ������ � �����
    Application.ActiveWindow.Page.Shapes(PRO).Cells("User.LineLenight").FormulaForceU = LenightCalc

End Sub

Sub CloneSectionUniverseNames(ByVal SectionIndex As Integer, ByVal ShapeFromID As Long, ByVal ShapeToID As Long)
'��������� ����������� ������������ ������ ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

On Error GoTo EX

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---��������� ������� ������ � ��������� SectionIndex � � ������ ���������� ������� �
    If (ShapeTo.SectionExists(SectionIndex, 0) = 0) And Not (ShapeFrom.SectionExists(SectionIndex, 0) = 0) Then
        'MsgBox "������� ����� ������"
        ShapeTo.AddSection (SectionIndex)
    End If

'On Error Resume Next
'---��������� ���� ������ �� �������� ����-�����
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            'If (ShapeTo.RowExists(SectionIndex, RowNum, 0) = 0) And Not (ShapeFrom.RowExists(SectionIndex, RowNum, 0) = 0) Then
                'MsgBox "create"
                ShapeTo.AddRow SectionIndex, RowNum, 0
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).RowNameU = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).RowName
            'End If
        Next RowNum
        
    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionUniverseNames"
    
End Sub

Sub CloneSectionScratchNames(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� ������������ ������ ������� ������ Scratch �� ������(ShapeFrom) � ������(ShapeTo)

On Error GoTo EX

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---��������� ������� ������ � ��������� SectionIndex � � ������ ���������� ������� �
    If (ShapeTo.SectionExists(visSectionScratch, 0) = 0) And Not (ShapeFrom.SectionExists(visSectionScratch, 0) = 0) Then
        ShapeTo.AddSection (visSectionScratch)
    End If

'---��������� ���� ������ �� �������� ����-�����
    For RowNum = 0 To ShapeFrom.RowCount(visSectionScratch) - 1
            ShapeTo.AddRow visSectionScratch, RowNum, 0
    Next RowNum
            
    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionScratchNames"
    
End Sub

Sub CloneSectionUniverseValues(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'---��������� ����������� ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

On Error GoTo 10

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---��������� ���� ������ �� �������� ����-�����
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            
        '---��������� ���� ������ � �������� � ������
            For CellNum = 0 To ShapeFrom.RowsCellCount(SectionIndex, RowNum) - 1
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Formula = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).Formula
                'MsgBox RowNum & ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Name
            Next CellNum
        Next RowNum

    Exit Sub
10:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionUniverseValues"

End Sub

Sub CloneSectionScratchValues(ShapeFromID As Long, ShapeToID As Long)
'---��������� ����������� ������� ������ Scratch �� ������(ShapeFrom) � ������(ShapeTo)

On Error GoTo EX

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
    Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---��������� ���� ������ �� �������� ����-�����
        For RowNum = 0 To ShapeFrom.RowCount(visSectionScratch) - 1
            
        '---��������� ���� ������ � �������� � ������
            For CellNum = 0 To ShapeFrom.RowsCellCount(visSectionScratch, RowNum) - 1
                ShapeTo.CellsSRC(visSectionScratch, RowNum, CellNum).Formula = _
                    ShapeFrom.CellsSRC(visSectionScratch, RowNum, CellNum).Formula
            Next CellNum
        Next RowNum

    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "Document_DocumentOpened"

End Sub

Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'��������� ������������ ������ �� ������ "Line"

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ���������� ���������� �����
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowLine, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowLine, j).Formula
    Next j

End Sub

Sub CloneSecEvent(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Event"
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j

End Sub

Sub CloneSecMiscellanious(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Miscellanious" - ���������� ������ Comment
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
'Dim RowNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---����������� ���� Comment ����� ������ �������� �� ������ �������.
    ShapeTo.CellsSRC(visSectionObject, visRowMisc, visComment).Formula = _
        ShapeFrom.CellsSRC(visSectionObject, visRowMisc, visComment).Formula

'---������� ��������� ����������
Set ShapeFrom = Nothing
Set ShapeTo = Nothing

End Sub


Sub ImportHoseInformation()
'��������� ��� ������� ������� ������ ������ (�������� �����)

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
    '---��������� ������ �� ����� ���� ������
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "�� ������� �� ���� ������!", vbInformation
        Exit Sub
    End If
    
    '---���������, �� �������� �� ��������� ������ ��� ������� ��� ������ ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � �������� �����", vbInformation
        Exit Sub
    End If
    
    '---��������� �������� �� ��������� ������ ������
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        MsgBox "��������� ������ �� �������� ������!", vbInformation
        Exit Sub
    End If

    '---��������� ������� �� � ������ �������� ������ ������� User.GFS_Aspect, ���� ���, �� ������� ��
    If Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) = False Then
        If Application.ActivePage.PageSheet.SectionExists(visSectionUser, 0) = False Then
            Application.ActivePage.PageSheet.AddSection (visSectionUser)
        End If
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", visTagDefault
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").Formula = 1
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect.Prompt").FormulaU = """������"""
    End If

'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("�����.vss").Masters("����� - ������").Shapes(1)
    IDTo = ShapeTo.ID   'Application.ActivePage.Shapes("Sheet.2").ID
    'IDFrom = ShapeFrom.Index
    IDFrom = Application.Documents("�����.vss").Masters("����� - ������").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User, Prop, Action, Controls
    CloneSectionUniverseNames 240, IDFrom, IDTo
    CloneSectionUniverseNames 242, IDFrom, IDTo
    CloneSectionUniverseNames 243, IDFrom, IDTo
    CloneSectionScratchNames IDFrom, IDTo

'---�������� ������� ����� ��� ��������� ������
    CloneSectionUniverseValues 240, IDFrom, IDTo
    CloneSectionUniverseValues 242, IDFrom, IDTo
    CloneSectionUniverseValues 243, IDFrom, IDTo
    CloneSectionScratchValues IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSectionLine IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---������������ ���������� �������� �����
    ReconnectHose ShapeTo

''---��������� ���� ������� ���������� ������
'On Error Resume Next
'Application.DoCmd (1312)

    LenightSetInner (ShapeTo.Name)

'---������� ��������� ����������
    Set ShapeTo = Nothing
    Set ShapeFrom = Nothing
    
Exit Sub
EX:
    '---������� ��������� ����������
    Set ShapeTo = Nothing
    Set ShapeFrom = Nothing
    SaveLog Err, "ImportHoseInformation"
End Sub


Sub ImportVHoseInformation()
'��������� ��� ������� ������� ������ ������ (����������� �����)

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
    '---��������� ������ �� ����� ���� ������
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "�� ������� �� ���� ������!", vbInformation
        Exit Sub
    End If
    
    '---���������, �� �������� �� ��������� ������ ��� ������� ��� ������ ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � �������� �����", vbInformation
        Exit Sub
    End If
    
    '---��������� �������� �� ��������� ������ ������
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        MsgBox "��������� ������ �� �������� ������!", vbInformation
        Exit Sub
    End If


'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("�����.vss").Masters("����������� �����").Shapes(1)
    IDTo = ShapeTo.ID
    'IDFrom = ShapeFrom.Index  'Application.Documents("����.vss").Masters("������� �������������").Index
    IDFrom = Application.Documents("�����.vss").Masters("����������� �����").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User, Prop, Action, Controls
    CloneSectionUniverseNames 240, IDFrom, IDTo
    CloneSectionUniverseNames 242, IDFrom, IDTo
    CloneSectionUniverseNames 243, IDFrom, IDTo
    CloneSectionScratchNames IDFrom, IDTo

'---�������� ������� ����� ��� ��������� ������
    CloneSectionUniverseValues 240, IDFrom, IDTo
    CloneSectionUniverseValues 242, IDFrom, IDTo
    CloneSectionUniverseValues 243, IDFrom, IDTo
    CloneSectionScratchValues IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSectionLine IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---������������ ���������� �������� �����
    ReconnectHose ShapeTo

'---��������� ���� ������� ���������� ������
    'Application.DoCmd (1312)

    LenightSetInner (ShapeTo.Name)
    
Exit Sub
EX:
    SaveLog Err, "ImportVHoseInformation"
End Sub


'--------------------------------��������� ������ ����--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'������� ���������� ����� ���� � ������� ��������� ���������������� ���� � ��������� �������
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, LayerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.Layer

    On erro GoTo EX
'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = Application.Documents("�����.vss").Masters(ShapeFromID).Shapes(1)
'    MsgBox Application.Documents("�����.vss").Masters(ShapeFromID)
    
'---�������� �������� ���� �������������� ������ � �������� ���������
    LayerName = ShapeFrom.Layer(1).Name

'---��������� ���� �� � ������� ��������� ���� � ����� ������
    For i = 1 To Application.ActivePage.Layers.Count
        If Application.ActivePage.Layers(i).Name = LayerName Then
            Flag = True
        End If
    Next i

'---� ������������ � ���������� ��������� ���������� ����� ���� � ������� ���������
    If Flag = True Then
        LayerNumber = Application.ActivePage.Layers(LayerName).Index
    Else
    '---������� ����� ���� � ������ ���� � �������� ����������� �������� ������
        Set vsoLayer = Application.ActiveWindow.Page.Layers.Add(LayerName)
        vsoLayer.NameU = LayerName
        vsoLayer.CellsC(visLayerColor).FormulaU = "255"
        vsoLayer.CellsC(visLayerStatus).FormulaU = "0"
        vsoLayer.CellsC(visLayerVisible).FormulaU = "1"
        vsoLayer.CellsC(visLayerPrint).FormulaU = "1"
        vsoLayer.CellsC(visLayerActive).FormulaU = "0"
        vsoLayer.CellsC(visLayerLock).FormulaU = "0"
        vsoLayer.CellsC(visLayerSnap).FormulaU = "1"
        vsoLayer.CellsC(visLayerGlue).FormulaU = "1"
        vsoLayer.CellsC(visLayerColorTrans).FormulaU = "0%"
    '---����������� ����� ������ ����
        LayerNumber = Application.ActivePage.Layers(LayerName).Index
    End If
        
LayerImport = Chr(34) & LayerNumber - 1 & Chr(34)
Exit Function
EX:
    SaveLog Err, "LayerImport"
End Function


'-----------------------���������� �������� ����� � ���----------------------------------------------------------
Private Sub ReconnectHose(ByRef ShpObj As Visio.Shape)
'��������� ��������� ��������� ���������� ��� ������ ������ �������� �����
Dim C_ConnectionsTrace As c_HoseConnector
Dim vO_Conn As Visio.Connect

    '---������� ��������� ������ c_HoseConnector ��� ������������� ���������� �����
    Set C_ConnectionsTrace = New c_HoseConnector
    '---��� ���� ���������� ���������� � �������� ����� ��������� ��������� ����������
    For Each vO_Conn In ShpObj.Connects
        C_ConnectionsTrace.Ps_ConnectionAdd vO_Conn
    Next vO_Conn

'---������� ��������� ����������
Set C_ConnectionsTraceLoc = Nothing
End Sub



'-----------------------��������� ��������� �����----------------------------------------------------------
'Public Sub MakeHoseLine()
''����� ��������� � �������� �����
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
''    On Error GoTo Tail
'
''---��������� ��������� ������� �����������, �������� ������ � ����� �������� ��������� �������
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---�������������� �������� ������
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---�������� ������ ��� ������
'    '---��������� ��������� ��������� ������ �������������
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
'    '---��������� ��������� ��������� ������� ���������� �������
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("�_������", "�������� ������")
'    '---��������� ��������� ��������� ������� ��������� �������
'    HoseDiametersListImport (ShpInd)
'    '---��������� ��������� ��������� �������� ������������� �������
'    HoseResistanceValueImport (ShpInd)
'    '---��������� ��������� ��������� �������� ���������� ����������� �������
'    HoseMaxFlowValueImport (ShpInd)
'    '---��������� ��������� ��������� �������� ����� �������
'    HoseWeightValueImport (ShpInd)
'
''---������������� �������� �������� ������� ��� ������
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---�������� ������
''    Ctrl.State = False
'
''---��������� ���� ������� ���������� ������
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---������� �� ��������� ���������
'    Application.EventsEnabled = True
'End Sub

Public Sub MakeVHoseLine()
'����� ��������� �� ����������� �������� �����
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail

'---��������� ��������� ������� �����������, �������� ������ � ����� �������� ��������� �������
    Application.EventsEnabled = False
    ImportVHoseInformation
    Application.EventsEnabled = True
    
'---�������������� �������� ������
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---�������� ������ ��� ������
    '---��������� ��������� ��������� ������ �������������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

'---������������� �������� �������� ������� ��� ������
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
    
'---��������� ���� ������� ���������� ������
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
    Application.EventsEnabled = True
End Sub

Public Sub MakeNapVsasHoseLine()
'����� ��������� � �������-����������� �������� �����
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail

'---��������� ��������� ������� �����������, �������� ������ � ����� �������� ��������� �������
    Application.EventsEnabled = False
    ImportVHoseInformation
    Application.EventsEnabled = True
    
'---�������������� �������� ������
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---�������� ������ ��� ������
    '---��������� ��������� ��������� ������ �������������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")

'---������������ �������� ��� �������-����������� �����
    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(1,Prop.LineType.Format)"
    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(0,Prop.HoseDiameter.Format)"

'---������������� �������� �������� ������� ��� ������
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
    
'---��������� ���� ������� ���������� ������
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
    Application.EventsEnabled = True
End Sub

Public Sub MakeHoseLine(ByVal hoseDiameterIndex As Integer, ByVal lineType As Byte)
'����� ��������� � �������� ����� � ��������� �����������
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer
Dim diameter As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail

'---��������� ��������� ������� �����������, �������� ������ � ����� �������� ��������� �������
    Application.EventsEnabled = False
    ImportHoseInformation
    Application.EventsEnabled = True
    
'---�������������� �������� ������
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---�������� ������ ��� ������
    '---��������� ��������� ��������� ������ �������������
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
    '---��������� ��������� ��������� ������� ���������� �������
    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("�_������", "�������� ������")
    '---��������� ��������� ��������� ������� ��������� �������
    HoseDiametersListImport (ShpInd)
    '---��������� ��������� ��������� �������� ������������� �������
    HoseResistanceValueImport (ShpInd)
    '---��������� ��������� ��������� �������� ���������� ����������� �������
    HoseMaxFlowValueImport (ShpInd)
    '---��������� ��������� ��������� �������� ����� �������
    HoseWeightValueImport (ShpInd)
        
'---������������ �������� ��� ������������� �����
    diameter = Index(hoseDiameterIndex, ShpObj.Cells("Prop.HoseDiameter.Format").Formula, ";")
    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(" & diameter & ",Prop.HoseDiameter.Format)"
    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(" & lineType & ",Prop.LineType.Format)"
        
'---������������� �������� �������� ������� ��� ������
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
        
'---��������� ���� ������� ���������� ������
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
    Application.EventsEnabled = True
End Sub

'Public Sub MakeShortHoseLine()
''����� ��������� � �������� ����� �� �������������� �������� �����
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
'    On Error GoTo Tail
'
''---��������� ��������� ������� �����������, �������� ������ � ����� �������� ��������� �������
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---�������������� �������� ������
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---�������� ������ ��� ������
'    '---��������� ��������� ��������� ������ �������������
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("�������������", "�������������")
'    '---��������� ��������� ��������� ������� ���������� �������
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("�_������", "�������� ������")
'    '---��������� ��������� ��������� ������� ��������� �������
'    HoseDiametersListImport (ShpInd)
'    '---��������� ��������� ��������� �������� ������������� �������
'    HoseResistanceValueImport (ShpInd)
'    '---��������� ��������� ��������� �������� ���������� ����������� �������
'    HoseMaxFlowValueImport (ShpInd)
'    '---��������� ��������� ��������� �������� ����� �������
'    HoseWeightValueImport (ShpInd)
'
''---������������ �������� ��� ������������� �����
'    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(2,Prop.HoseDiameter.Format)"
'    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(2,Prop.LineType.Format)"
'
''---������������� �������� �������� ������� ��� ������
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---��������� ���� ������� ���������� ������
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---������� �� ��������� ���������
'    Application.EventsEnabled = True
'End Sub

