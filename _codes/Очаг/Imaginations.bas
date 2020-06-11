Attribute VB_Name = "Imaginations"

Sub SquareSetInner(shpObjName As String) '����������
'��������� ���������� ���������� ���� ���������� ������ �������� ������� ������
Dim SquareCalc As Integer
Dim ShpObj As Visio.Shape

'---���������� ��������� ���������� ��� �������� ������
    Set ShpObj = Application.ActivePage.Shapes(shpObjName)

'---���������� � �������
    SquareCalc = ShpObj.AreaIU * 0.00064516 '��������� �� ���������� ������ � ���������� �����
    ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub

Sub CloneSectionUniverseNames(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� ������������ ������ ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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
            
End Sub

Sub CloneSectionUniverseValues(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'---��������� ����������� ������� ��������� ������ �� ������(ShapeFrom) � ������(ShapeTo)

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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


'---------------------------------��������� � ������� �������-------------------------------------
Sub ImportAreaInformation() '(������� ������)
'��������� ��� ������� ������� ������� ������

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

''---��������� ������ �� ����� ���� ������
    'If Application.ActiveWindow.Selection.Count < 1 Then
    '    MsgBox "�� ������� �� ���� ������!", vbInformation
    '    Exit Sub
    'End If
    '
'---���������, ��� ������ ����� ���� ������
    Debug.Print Application.ActiveWindow.Selection(1).Name
    If Application.ActiveWindow.Selection(1).CellExists("User.visObjectType", 0) Then
        If Application.ActiveWindow.Selection(1).Cells("User.visObjectType") = 104 Then
            PF_GeometryCopy Application.ActiveWindow.Selection(1)
        End If
    End If
'
''---���������, �� �������� �� ��������� ������ ��� �������� ��� ������ ������� � ������������ ����������
    'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
    '    MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ���� �������", vbInformation
    '    Exit Sub
    'End If
'
''---��������� �������� �� ��������� ������ ��������
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "��������� ������ �� ����� �������!", vbInformation
    '    Exit Sub
    'End If


'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("������� �������������").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = ThisDocument.Masters("������� �������������").Index

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
    CloneSecFill IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo
    CreateTextFild IDTo

'---������������� �������� �������� ������� �� ������ TheDoc!User.CurrentTime
    ShapeTo.Cells("Prop.SquareTime").Formula = _
        "=DATETIME(" & CStr(Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---��������� ���� ������� ���������� ������
    On Error Resume Next '�� ������ ��� ������� ���������� �����
    'Application.DoCmd (1312)
    If VfB_NotShowPropertiesWindow = False Then Application.DoCmd (1312) '� ������ ���� ����� ���� �������, ���������� ����

    SquareSetInner (ShapeTo.Name)

End Sub

Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'��������� ������������ ������ �� ������ "Line"

'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)
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
Dim RowNum As Integer   ', CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j

End Sub

Sub CloneSecFill(ShapeFromID As Long, ShapeToID As Long)
'��������� ����������� �������� ����� ��� ������ "Fill"
'---��������� ����������
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---������� ������ ����� � ����� �������, � � ������ ���������� � ��� ������� - ������� ��.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
    Next j

End Sub


Sub CreateTextFild(ShapeToID As Long)
'��������� ���������� ����
'---��������� ����������
Dim ShapeTo As Visio.Shape
Dim vsoCharacters As Visio.Characters

'---����������� �������� ����������
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)
Set vsoCharacters = ShapeTo.Characters

'MsgBox ShapeTo.Name

'---������� ����� ��������� ���� � ����������� ��� ��������
    vsoCharacters.Begin = 0
    vsoCharacters.End = 0
    vsoCharacters.AddCustomFieldU "GUARD(Prop.FireCategorie&Prop.IntenseShowType&Prop.FireDescription)", visFmtNumGenNoUnits
    ShapeTo.CellsSRC(visSectionCharacter, 0, visCharacterLangID).FormulaU = 1033

'---�������� ����� ������ �� ��������� � �������
    ShapeTo.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 1
    ShapeTo.CellsSRC(visSectionObject, visRowMisc, visHideText).FormulaU = True

'---������� ����������
Set ShapeTo = Nothing
Set vsoCharacters = Nothing

End Sub


'---------------------------------��������� � �����-------------------------------------
Sub ImportStormInformation()
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
''---���������, �� �������� �� ��������� ������ ��� ������� ��� ������ ������� � ������������ ����������
    'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
    '    MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � �������� �����", vbInformation
    '    Exit Sub
    'End If
    '
''---��������� �������� �� ��������� ������ ��������
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "��������� ������ �� ����� ������� � �� ����� ���� �������� � �������� �����!", vbInformation
    '    Exit Sub
    'End If

'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("�������� �����").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = ThisDocument.Masters("�������� �����").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User � Actions
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 240, IDFrom, IDTo  'Action

'---�������� ������� ����� ��� ��������� ������
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 240, IDFrom, IDTo  'Action

'---�������� ������� ����� ��� ��������� ������
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

End Sub


'---------------------------------��������� � ����������� ����-------------------------------------
Sub ImportFogInformation()
'��������� ��� ������� ������� ������ ������

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
'---��������� ������ �� ����� ���� ������
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "�� ������� �� ���� ������!", vbInformation
        Exit Sub
    End If

'---���������, �� �������� �� ��������� ������ ��� �������� ��� ������ ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ����������� ����", vbInformation
        Exit Sub
    End If

'---��������� �������� �� ��������� ������ ��������
    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
        MsgBox "��������� ������ �� ����� ������� � �� ����� ���� ��������!", vbInformation
        Exit Sub
    End If

'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("����������").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = ThisDocument.Masters("����������").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User, Prop, Action, Controls
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---�������� ������� ����� ��� ��������� ������
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

Exit Sub
EX:
    SaveLog Err, "ImportFogInformation"
End Sub


'---------------------------------��������� � ���� ���������-------------------------------------
Sub ImportRushInformation()
'��������� ��� ������� ������� ������ ������

'---��������� ����������
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
'---��������� ������ �� ����� ���� ������
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "�� ������� �� ���� ������!", vbInformation
        Exit Sub
    End If

'---���������, �� �������� �� ��������� ������ ��� �������� ��� ������ ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ���� ���������", vbInformation
        Exit Sub
    End If

'---��������� �������� �� ��������� ������ ��������
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "��������� ������ �� ����� ������� � �� ����� ���� ��������!", vbInformation
    '    Exit Sub
    'End If

'---����������� ���������� ������� �����(ShapeFrom � ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("���� ���������").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = ThisDocument.Masters("���� ���������").Index

'---������� ����������� ����� ���������������� ����� ��� ������ User, Prop, Action, Controls
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---�������� ������� ����� ��� ��������� ������
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---��������� ������� ����� (��� ������)
    ShapeTo.Cells("Prop.RushTime").Formula = "DateTime(" & ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate) & ")"

'---����������� ����� ����
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---��������� ���� ������� ���������� ������
    On Error Resume Next '�� ������ ��� ������� ���������� �����

    SquareSetInner (ShapeTo.Name)
    Application.DoCmd (1312)
    
Exit Sub
EX:
    SaveLog Err, "ImportRushInformation"
End Sub

'--------------------------------��������� ������ ����--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'������� ���������� ����� ���� � ������� ��������� ���������������� ���� � ��������� �������
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, layerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.layer

'---����������� ��������� ���������� ������(ShapeFrom � ShpeTo) � ������������ � ���������
    Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)

'---�������� �������� ���� �������������� ������ � �������� ���������
    layerName = ShapeFrom.layer(1).Name

'---��������� ���� �� � ������� ��������� ���� � ����� ������
    For i = 1 To Application.ActivePage.Layers.Count
        If Application.ActivePage.Layers(i).Name = layerName Then
            Flag = True
        End If
    Next i

'---� ������������ � ���������� ��������� ���������� ����� ���� � ������� ���������
    If Flag = True Then
        LayerNumber = Application.ActivePage.Layers(layerName).Index
    Else
    '---������� ����� ���� � ������ ���� � �������� ����������� �������� ������
        Set vsoLayer = Application.ActiveWindow.Page.Layers.Add(layerName)
        vsoLayer.NameU = layerName
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
        LayerNumber = Application.ActivePage.Layers(layerName).Index
    End If
        
LayerImport = Chr(34) & LayerNumber - 1 & Chr(34)

End Function



'-------------------------------------��������� ������ ����� � ������� ������--------------------------------------
Public Function PF_GeometryCopy(ByRef OriginalShp As Visio.Shape) As Visio.Shape
'����� ������� ����������� ������ � ������ ��������� �������� ������
'Dim OriginalShp As Visio.Shape
Dim ReplicaShape As Visio.Shape
Dim i As Integer
Dim j As Integer
Dim k As Integer

    On Error GoTo EX
'    Set OriginalShp = Application.ActiveWindow.Page.Shapes.ItemFromID(231)
    Set ReplicaShape = Application.ActiveWindow.Page.DrawRectangle(0, 0, 100, 100)
    
    '---��������� ���������
    i = 0
    Do While OriginalShp.SectionExists(visSectionFirstComponent + i, 0)
        If i > 0 Then
            ReplicaShape.AddSection (visSectionFirstComponent + i)
            ReplicaShape.AddRow visSectionFirstComponent + i, visRowComponent, visTagComponent
        Else
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
        End If
        
        j = 1
        Do While OriginalShp.RowExists(visSectionFirstComponent + i, j, 0)
            ReplicaShape.AddRow visSectionFirstComponent + i, j, OriginalShp.RowType(visSectionFirstComponent + i, j)
            
            k = 0
            Do While OriginalShp.CellsSRCExists(visSectionFirstComponent + i, j, k, 0)
                ReplicaShape.CellsSRC(visSectionFirstComponent + i, j, k).FormulaU = _
                    OriginalShp.CellsSRC(visSectionFirstComponent + i, j, k).FormulaU
                
                k = k + 1
            Loop
            j = j + 1
        Loop
        i = i + 1
    Loop
    
    '---������������ ��������� � ������� �������� ������ � ������ �������
        ReplicaShape.Cells("Width").FormulaU = OriginalShp.Cells("Width").FormulaU
        ReplicaShape.Cells("Height").FormulaU = OriginalShp.Cells("Height").FormulaU
        ReplicaShape.Cells("LocPinX").FormulaU = OriginalShp.Cells("LocPinX").FormulaU
        ReplicaShape.Cells("LocPinY").FormulaU = OriginalShp.Cells("LocPinY").FormulaU
        ReplicaShape.Cells("PinX").FormulaU = OriginalShp.Cells("PinX").FormulaU
        ReplicaShape.Cells("PinY").FormulaU = OriginalShp.Cells("PinY").FormulaU
        
Set PF_GeometryCopy = ReplicaShape

Exit Function
EX:
    MsgBox "�������� �������������� ������! ���� ��� ����� ����������� - ���������� � ������������", , ThisDocument.Name
    SaveLog Err, "Document_DocumentOpened"
End Function








