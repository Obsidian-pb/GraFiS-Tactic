Attribute VB_Name = "m_WorkWithStyles"
'----------------------------� ������ �������� ��������� ������ �� �������--------------------------------------
Dim WaterSource_StyleSet(7) As String


Private Sub s_StyleSetsDeclare()
'��������� ������ �������� ������ ��� ����������
'---����� "�������������"
    WaterSource_StyleSet(0) = "��_������"
    WaterSource_StyleSet(1) = "��_�����"
    WaterSource_StyleSet(2) = "��_�������"
    WaterSource_StyleSet(3) = "��_������1"
    WaterSource_StyleSet(4) = "��_��"
    WaterSource_StyleSet(5) = "��_�������"
    WaterSource_StyleSet(6) = "��_�������"
    
End Sub


Public Sub StyleExport()
'������� ��������� �������� ������ � �������� ��������
Dim vO_Doc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer

'---������� ����� �������� ������
    s_StyleSetsDeclare

'---��������� ����� ��������� ���������
    Set vO_Doc = Application.ActiveDocument
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(WaterSource_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(WaterSource_StyleSet(i))
        '---��������� ���� �� �������� ����� WaterSource_StyleSet(i) � �������� ���������
            If StyleExist(WaterSource_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh WaterSource_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� WaterSource_StyleSet(i) � �������� ���������
                vO_Doc.Drop vO_Stl, 0, 0
            End If
        Next i

'---������� �������
'    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing

End Sub


Private Sub StyleRefresh(as_StyleName As String)
'��������� ���������� ����� as_StyleName � �������� ���������
Dim vO_StyleFrom As Visio.style
Dim vO_StyleTo As Visio.style
Dim vs_RowName As String

'---������� ����������� ����� ��������
    Set vO_StyleFrom = ThisDocument.Styles(as_StyleName)
    Set vO_StyleTo = ActiveDocument.Styles(as_StyleName)

'---��������� ������ ��� ����� "�����"
    If vO_StyleFrom.Cells("EnableTextProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("Char.Font").FormulaU = vO_StyleFrom.Cells("Char.Font").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Size").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Size").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.FontScale").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.FontScale").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Letterspace").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Letterspace").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Color").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Color").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.ColorTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.ColorTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Style").FormulaU = vO_StyleFrom.Cells("Char.Style").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Case").FormulaU = vO_StyleFrom.Cells("Char.Case").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Pos").FormulaU = vO_StyleFrom.Cells("Char.Pos").ResultStr(visUnitsString)
    End If
    
'---��������� ������ ��� ����� "�����"
    If vO_StyleFrom.Cells("EnableLineProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("LinePattern").FormulaU = vO_StyleFrom.Cells("LinePattern").ResultStr(visUnitsString)
'        vO_StyleTo.Cells("LineWeight").FormulaU = vO_StyleFrom.Cells("LineWeight").ResultStr(visUnitsString)
        vO_StyleTo.Cells("LineColor").FormulaU = Chr(34) & vO_StyleFrom.Cells("LineColor").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("LineCap").FormulaU = vO_StyleFrom.Cells("LineCap").ResultStr(visUnitsString)
        vO_StyleTo.Cells("LineColorTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("LineColorTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Rounding").FormulaU = Chr(34) & vO_StyleFrom.Cells("Rounding").ResultStr(visUnitsString) & Chr(34)
    End If
    
'---��������� ������ ��� ����� "�������"
    If vO_StyleFrom.Cells("EnableFillProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("FillForegnd").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillForegnd").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillForegndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillForegndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillBkgnd").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillBkgnd").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillBkgndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillBkgndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillPattern").FormulaU = vO_StyleFrom.Cells("FillPattern").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwForegnd").FormulaU = vO_StyleFrom.Cells("ShdwForegnd").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwForegndTRans").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShdwForegndTRans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShdwBkgnd").FormulaU = vO_StyleFrom.Cells("ShdwBkgnd").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwBkgndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShdwBkgndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShdwPattern").FormulaU = vO_StyleFrom.Cells("ShdwPattern").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShapeShdwOffsetX").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwOffsetX").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwOffsetY").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwOffsetY").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwType").FormulaU = vO_StyleFrom.Cells("ShapeShdwType").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShapeShdwObliqueAngle").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwObliqueAngle").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwScaleFactor").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwScaleFactor").ResultStr(visUnitsString) & Chr(34)
    End If
    

'---������� �������
    Set vO_StyleFrom = Nothing
    Set vO_StyleTo = Nothing
    Set vO_Stenc = Nothing
End Sub


Private Function StyleExist(as_StyleName As String) As Boolean
'������� ���������� ������ �������� ������� ������� ����� � �������� ���������
Dim vO_Style As Visio.style

    StyleExist = False
    
    For Each vO_Style In Application.ActiveDocument.Styles
        If vO_Style.Name = as_StyleName Then StyleExist = True
    Next vO_Style

End Function


