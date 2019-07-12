Attribute VB_Name = "Imaginations"

'---------------------------------��������� � �������� ������������-------------------------------------
Public Sub ImportOpenWaterInformation()
'��������� ��� ������� ������� ������ ������

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

''---��������� ������ �� ����� ���� ������
'If Application.ActiveWindow.Selection.Count < 1 Then
'    MsgBox "�� ������� �� ���� ������!", vbInformation
'    Exit Sub
'End If
'
''---���������, �� �������� �� ��������� ������ ��� �������� ��� ������ ������� � ������������ ����������
'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
'    MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ������������ ������", vbInformation
'    Exit Sub
'End If
'
''---��������� �������� �� ��������� ������ ��������
'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
'    MsgBox "��������� ������ �� ����� ������� � �� ����� ���� �������� � ������������ ������!", vbInformation
'    Exit Sub
'End If


'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
Set ShapeTo = Application.ActiveWindow.Selection(1)
Set ShapeFrom = Application.Documents("�������������.vss").Masters("�������� ������������").Shapes(1)
IDTo = ShapeTo.ID
IDFrom = Application.Documents("�������������.vss").Masters("�������� ������������").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User, Prop, Action, Controls
CloneSectionUniverseNames 240, IDFrom, IDTo  'Action
CloneSectionUniverseNames 242, IDFrom, IDTo  'User
CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---�������� ������� ����� ��� ��������� ������
CloneSectionUniverseValues 240, IDFrom, IDTo  'Action
CloneSectionUniverseValues 242, IDFrom, IDTo  'User
CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
CloneSecEvent IDFrom, IDTo
CloneSectionLine IDFrom, IDTo
CloneSecMiscellanious IDFrom, IDTo
CloneSecFill IDFrom, IDTo
'CreateTextFild IDTo

'---����������� ����� ����
ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---��������� ���� ������� ���������� ������
On Error Resume Next '�� ������ ��� ������� ���������� �����
Application.DoCmd (1312)

'---������� �����
Set ShapeTo = Nothing
Set ShapeFrom = Nothing

End Sub

Private Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'��������� ������������ ������ �� ������ "Line"

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ���������� ���������� �����
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowLine, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowLine, j).Formula
    Next j

End Sub

Private Sub CloneSecEvent(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Event"
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

    On erroro GoTo EX

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j
Exit Sub
EX:
    SaveLog Err, "CloneSecEvent"
End Sub

Private Sub CloneSecFill(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Fill"
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowFill, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowFill, j).Formula
    Next j

End Sub

Sub CloneSecMiscellanious(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Miscellanious"
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
    Next j

End Sub




Sub CloneSectionUniverseNames(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� ������������ ������ ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'MsgBox Application.Documents("�������������.vss").Masters(ShapeFromID).Name

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
            
End Sub

Sub CloneSectionUniverseValues(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'---��������� ����������� ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---��������� ���� ������ �� �������� ����-�����
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            
        '---��������� ���� ������ � �������� � ������
            For CellNum = 0 To ShapeFrom.RowsCellCount(SectionIndex, RowNum) - 1
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Formula = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).Formula
                'MsgBox RowNum & ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Name
            Next CellNum
        Next RowNum


End Sub




'--------------------------------��������� ������ ����--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'������� ���������� ����� ���� � ������� ��������� ���������������� ���� � ��������� �������
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, LayerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.Layer

    On Error GoTo EX
'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = Application.Documents("�������������.vss").Masters(ShapeFromID).Shapes(1)

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

