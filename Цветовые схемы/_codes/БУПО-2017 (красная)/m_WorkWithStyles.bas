Attribute VB_Name = "m_WorkWithStyles"
Option Explicit
'----------------------------� ������ �������� ��������� ������ �� �������--------------------------------------
'----------------------------������ ������ ���������------------------------------------------------------------


Public Sub StyleExport()
'������� ��������� ���������� ������
'---��������� ����� "�������� �������"
    If StancilExist("�������� �������.vss") Then Refresh_PT
'---��������� ����� "������� ������"
    If StancilExist("������� ������.vss") Then Refresh_PTO
'---��������� ����� "���"
    If StancilExist("���.vss") Then Refresh_PTV
'---��������� ����� "����"
    If StancilExist("����.vss") Then Refresh_GDZ
'---��������� ����� "�����"
    If StancilExist("�����.vss") Then Refresh_Line
'---��������� ����� "����� � ���������"
    If StancilExist("����� � ���������.vss") Then Refresh_RL
'---��������� ����� "�������������"
    If StancilExist("�������������.vss") Then Refresh_WaterSource
'---��������� ����� "����"
    If StancilExist("����.vss") Then Refresh_Fire
'---��������� ����� "���������� ���"
    If StancilExist("���������� ���.vss") Then Refresh_Mngmnt
'---��������� ����� "������"
    If StancilExist("������.vss") Then Refresh_Other


End Sub


Private Sub Refresh_PT()
'������� ��������� �������� ������ � �������� "�������� �������"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PT_StyleSet(9) As String '�������� �������

    PT_StyleSet(0) = "�_������_������"
    PT_StyleSet(1) = "�_��_��������"
    PT_StyleSet(2) = "�_������_�����"
    PT_StyleSet(3) = "�_������_������"
    PT_StyleSet(4) = "�_��_����������"
    PT_StyleSet(5) = "�_��������"
    PT_StyleSet(6) = "�_��������"
'    PT_StyleSet(7) = "�_�����_������"
'    PT_StyleSet(8) = "�_�����_������"
'    PT_StyleSet(9) = "�_�������������"
    PT_StyleSet(7) = "�_����_������"
    PT_StyleSet(8) = "�_����������"


    '---��������� ��� ������ ��������
        DocOpenClose "�������� �������.vss", 1
    
    '---��������� ����� ��������� "�������� �������"
        Set vO_Stenc = Application.Documents("�������� �������.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(PT_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(PT_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("�������� �������.vss", PT_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "�������� �������.vss", PT_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "�������� �������.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_PTO()
'������� ��������� �������� ������ � �������� "������� ������"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PTO_StyleSet(10) As String '������� ������
    
    PTO_StyleSet(0) = "�_���������"
    PTO_StyleSet(1) = "�_������_�����"
    PTO_StyleSet(2) = "�_��������"
    PTO_StyleSet(3) = "�_��������"
    PTO_StyleSet(4) = "�_��������"
    PTO_StyleSet(5) = "�_�����_������"
    PTO_StyleSet(6) = "�_�����_������"
    PTO_StyleSet(7) = "�_�������������"
    PTO_StyleSet(8) = "�_������_����������"
    PTO_StyleSet(9) = "�_������_�����������"


    '---��������� ��� ������ ��������
        DocOpenClose "������� ������.vss", 1
        
    '---��������� ����� ��������� "�������� �������"
        Set vO_Stenc = Application.Documents("������� ������.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(PTO_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(PTO_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("������� ������.vss", PTO_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "������� ������.vss", PTO_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "������� ������.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_PTV()
'������� ��������� �������� ������ � �������� "���"
Dim vs_StencName As String
'Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PTV_StyleSet(10) As String '���

    PTV_StyleSet(0) = "���_��������"
    PTV_StyleSet(1) = "���_�����������"
    PTV_StyleSet(2) = "���_��������"
    PTV_StyleSet(3) = "���_��������_�������"
    PTV_StyleSet(4) = "���_��������"
    PTV_StyleSet(5) = "�_��������"
    PTV_StyleSet(6) = "�_��"
    PTV_StyleSet(7) = "���_��_������"
    PTV_StyleSet(8) = "���_��_����"
    PTV_StyleSet(9) = "���_�����"


    '---��������� ��� ������ ��������
        DocOpenClose "���.vss", 1
    
    '---��������� ����� ��������� "���"
        Set vO_Stenc = Application.Documents("���.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(PTV_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(PTV_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("���.vss", PTV_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "���.vss", PTV_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "���.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_GDZ()
'������� ��������� �������� ������ � �������� "����"
Dim vs_StencName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim GDZ_StyleSet(5) As String '����

    GDZ_StyleSet(0) = "���_��������"
    GDZ_StyleSet(1) = "���_��������_�������"
    GDZ_StyleSet(2) = "���_�����"
    GDZ_StyleSet(3) = "���_����"
    GDZ_StyleSet(4) = "�_��������"


    '---��������� ��� ������ ��������
        DocOpenClose "����.vss", 1
        
    '---��������� ����� ��������� "����"
        Set vO_Stenc = Application.Documents("����.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(GDZ_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(GDZ_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("����.vss", GDZ_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "����.vss", GDZ_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "����.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_Line()
'������� ��������� �������� ������ � �������� "�����"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Line_StyleSet(7) As String '�����

    Line_StyleSet(0) = "�_��"
    Line_StyleSet(1) = "�_���"
    Line_StyleSet(2) = "�_��"
    Line_StyleSet(3) = "�_�������"
    Line_StyleSet(4) = "�_����"
    Line_StyleSet(5) = "�_������"
    Line_StyleSet(6) = "�_��������"

    '---��������� ��� ������ ��������
        DocOpenClose "�����.vss", 1
        
    '---��������� ����� ��������� "�����"
        Set vO_Stenc = Application.Documents("�����.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(Line_StyleSet()) - 1
        '---�������� ��������� ����� ���������
            Set vO_Stl = ThisDocument.Styles(Line_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("�����.vss", Line_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "�����.vss", Line_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "�����.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_RL()
'������� ��������� �������� ������ � �������� "����� � ���������"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim RL_StyleSet(6) As String '����� � ���������

    RL_StyleSet(0) = "�_�������"
    RL_StyleSet(1) = "�_����"
    RL_StyleSet(2) = "�_���"
    RL_StyleSet(3) = "�_����"
    RL_StyleSet(4) = "�_�������"
    RL_StyleSet(5) = "�_��������"

    '---��������� ��� ������ ��������
        DocOpenClose "����� � ���������.vss", 1
    
    '---��������� ����� ��������� "����� � ���������"
        Set vO_Stenc = Application.Documents("����� � ���������.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(RL_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(RL_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("����� � ���������.vss", RL_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "����� � ���������.vss", RL_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "����� � ���������.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_WaterSource()
'������� ��������� �������� ������ � �������� "�������������"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim WaterSource_StyleSet(7) As String '�������������

    WaterSource_StyleSet(0) = "��_������"
    WaterSource_StyleSet(1) = "��_�����"
    WaterSource_StyleSet(2) = "��_�������"
    WaterSource_StyleSet(3) = "��_������1"
    WaterSource_StyleSet(4) = "��_��"
    WaterSource_StyleSet(5) = "��_�������"
    WaterSource_StyleSet(6) = "��_�������"


    '---��������� ��� ������ ��������
        DocOpenClose "�������������.vss", 1
    
    '---��������� ����� ��������� "�������������"
        Set vO_Stenc = Application.Documents("�������������.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(WaterSource_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(WaterSource_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("�������������.vss", WaterSource_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "�������������.vss", WaterSource_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "�������������.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Fire()
'������� ��������� �������� ������ � �������� "����"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Fire_StyleSet(10) As String '����

    Fire_StyleSet(0) = "�_����������"
    Fire_StyleSet(1) = "�_���������"
    Fire_StyleSet(2) = "�_�������"
    Fire_StyleSet(3) = "�_���������"
    Fire_StyleSet(4) = "�_����"
    Fire_StyleSet(5) = "�_�������"
    Fire_StyleSet(6) = "�_�����������"
    Fire_StyleSet(7) = "�_�����������"
    Fire_StyleSet(8) = "�_�����"
    Fire_StyleSet(9) = "�_�����"

    '---��������� ��� ������ ��������
        DocOpenClose "����.vss", 1
    
    '---��������� ����� ��������� "����"
        Set vO_Stenc = Application.Documents("����.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(Fire_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(Fire_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("����.vss", Fire_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "����.vss", Fire_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "����.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Mngmnt()
'������� ��������� �������� ������ � �������� "���������� ���"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Mngmnt_StyleSet(3) As String '���������� ���

    Mngmnt_StyleSet(0) = "�_���"
    Mngmnt_StyleSet(1) = "��_���"
    Mngmnt_StyleSet(2) = "��_����"

    '---��������� ��� ������ ��������
        DocOpenClose "���������� ���.vss", 1
    
    '---��������� ����� ��������� "���������� ���"
        Set vO_Stenc = Application.Documents("���������� ���.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(Mngmnt_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(Mngmnt_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("���������� ���.vss", Mngmnt_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "���������� ���.vss", Mngmnt_StyleSet(i)
            Else
            '---���� ��� - ���������� ���� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "���������� ���.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Other()
'������� ��������� �������� ������ � �������� "������"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Other_StyleSet(1) As String '������

    Other_StyleSet(0) = "��_�������"
'    Mngmnt_StyleSet(1) = "��_���"
'    Mngmnt_StyleSet(2) = "��_����"


    '---��������� ��� ������ ��������
        DocOpenClose "������.vss", 1
    
    '---��������� ����� ��������� "������"
        Set vO_Stenc = Application.Documents("������.vss")
    
    '---���������� ��� ����� ���������
        For i = 0 To UBound(Other_StyleSet()) - 1
        '---�������� ��������� ����� ���������
        Set vO_Stl = ThisDocument.Styles(Other_StyleSet(i))
        '---��������� ���� �� �������� ����� vs_StyleName � ��������� vs_StencName
            If StyleExist("������.vss", Other_StyleSet(i)) Then
            '---���� ���� - ��������� ���
                StyleRefresh "������.vss", Other_StyleSet(i)
            Else
            '---���� ��� - ���������� ����� vs_StyleName � �������� vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---��������� ��� ������ ��������
        DocOpenClose "������.vss", 0

'---������� �������
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

'-------------------------------------��������� ����������-------------------------------------------------------
Private Sub StyleRefresh(as_StnclName As String, as_StyleName As String)
'��������� ���������� ����� as_StyleName � ��������� as_StnclName
Dim vO_Stenc As Visio.Document
Dim vO_StyleFrom As Visio.style
Dim vO_StyleTo As Visio.style
Dim vs_RowName As String

'---������� ����������� ����� ��������
    Set vO_Stenc = Application.Documents(as_StnclName)
    Set vO_StyleFrom = ThisDocument.Styles(as_StyleName)
    Set vO_StyleTo = vO_Stenc.Styles(as_StyleName)

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


Private Sub DocOpenClose(asS_StencName As String, asb_OpenClose As Byte)
'��������� ��������-�������� ��������� ��� ���������.  0 - �������; 1 - �������
Dim doc As Visio.Document
Dim pth As String

'---�������������� ����������� ��������
Set doc = Documents(asS_StencName)
pth = doc.fullName

'---��������� ��� ��������� �������� � ����������� �� ���������� asb_OpenClose
    If asb_OpenClose = 0 Then
        If doc.ReadOnly = False Then
            doc.Close
            Application.Documents.OpenEx pth, visOpenRO + visOpenDocked
        End If
    Else
        If doc.ReadOnly = True Then
            doc.Close
            Application.Documents.OpenEx pth, visOpenRW + visOpenDocked
        End If
    End If

'---������� �������
Set doc = Nothing

End Sub


'-----------------------------------------------------�������--------------------------------------------------
Private Function StyleExist(asS_StencName As String, as_StyleName As String) As Boolean
'������� ���������� ������ �������� ������� ������� ����� � ��������� ���������
Dim vO_Style As Visio.style

StyleExist = False

For Each vO_Style In Application.Documents(asS_StencName).Styles
    If vO_Style.Name = as_StyleName Then StyleExist = True
Next vO_Style

End Function

Private Function StancilExist(asS_StencName As String) As Boolean
'������� ���������� ������ �������� ������� ������� ���������
Dim vO_Stencil As Visio.Document
StancilExist = False

For Each vO_Stencil In Application.Documents
    If vO_Stencil.Name = asS_StencName Then StancilExist = True
Next vO_Stencil

End Function


