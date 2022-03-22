Attribute VB_Name = "Vichislenia"
Option Explicit

Sub SquareSet(ShpObj As Visio.Shape)
'��������� ���������� ���������� ���� ���������� ������ �������� ������� ������
'������ ��� ����� ������� ������
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 '��������� �� ���������� ������ � ���������� �����
ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub

Sub USquareSet(ShpObj As Visio.Shape)
'��������� ���������� ���������� ���� ���������� ������ �������� ������� ������
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 '��������� �� ���������� ������ � ���������� �����
ShpObj.Cells("User.SquareP").FormulaForceU = SquareCalc

End Sub


Sub s_SetFireTime(ShpObj As Visio.Shape, Optional showDoCmd As Boolean = True)
'��������� ���������� ������ ��������� User.FireTime �������� ������� ���������� ��� ����������� ������ "����"
Dim vD_CurDateTime As Double

On Error Resume Next

'---��������� ������������ �� ������ ������ �������
    If IsFirstDrop(ShpObj) Then
        '---����������� �������� ������� ������������� ������ ������� ��������
            vD_CurDateTime = Now()
            ShpObj.Cells("Prop.FireTime").FormulaU = _
                "DATETIME(" & str(vD_CurDateTime) & ")"
        
        '---���������� ���� ������� ������
            If showDoCmd Then Application.DoCmd (1312)
            
        '---���� � ����-����� ��������� ����������� ������ "User.FireTime", ������� �
            If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
                Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
            End If
            
        '---��������� ����� ������ �� ���� ������ ������ � ���� ���� ���������
            Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
                "DATETIME(" & str(CDbl(ShpObj.Cells("Prop.FireTime").Result(visDate))) & ")"
    Else
        '---���������� ���� ������� ������
            If showDoCmd Then Application.DoCmd (1312)
    End If
    '---��������� � �������� ������� �������� ������ � �������� ������������
    AddPageTimeProps ShpObj

End Sub

Public Sub ShowTimesForm(ByRef shp As Visio.Shape)
    F_Times.ShowMe shp
End Sub













'------------------�������� ����� ������ � ��������----------------------
Public Sub AddPageTimeProps(ByRef shpFire As Visio.Shape)
Dim tmpRowInd As Integer
'Dim tmpRow As Visio.Row
Dim shp As Visio.Shape

    
'    On Error Resume Next
    
    Set shp = Application.ActivePage.PageSheet
    
    '����� �������������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� �������������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ������������� ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 20
        SetCellVal shp, "Prop.FireTime", CellVal(shpFire, "Prop.FireTime", visDate) ' GetVal(fir)

        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� �������������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireTime" & Chr(34) & ",TheDoc!User.FireTime)+DEPENDSON(TheDoc!User.FireTime)"  '����� !!! ������� ����������� ����� ������� - �������, ����� ����� ���������� �� ������ ������������ �������� ����� Scratch
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireTime" & Chr(34) & ", Prop.FireTime) + DEPENDSON(Prop.FireTime)"      '����� !!! ������� ����������� ����� ������� - �������, ����� ����� ���������� �� ������ ������������ �������� ����� Scratch

    '����� �����������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FindTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� �����������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ����������� ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FindTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 21
'        SetCellVal shp, "Prop.FindTime", GetVal(fin)
        SetCellVal shp, "Prop.FindTime", CellVal(shpFire, "Prop.FindTime", visDate)

        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� �����������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FindTime" & Chr(34) & ",TheDoc!User.FindTime)+DEPENDSON(TheDoc!User.FindTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FindTime" & Chr(34) & ", Prop.FindTime) + DEPENDSON(Prop.FindTime)"
        
    '����� ���������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "InfoTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ���������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ��������� � ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.InfoTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 22
'        SetCellVal shp, "Prop.InfoTime", GetVal(inf)
        SetCellVal shp, "Prop.InfoTime", CellVal(shpFire, "Prop.InfoTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ���������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.InfoTime" & Chr(34) & ",TheDoc!User.InfoTime)+DEPENDSON(TheDoc!User.InfoTime)"  '
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.InfoTime" & Chr(34) & ", Prop.InfoTime) + DEPENDSON(Prop.InfoTime)"

    '����� �������� ������� �������������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstArrivalTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ��������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� �������� � ����� ������ ������� �������������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FirstArrivalTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 23
'        SetCellVal shp, "Prop.FirstArrivalTime", GetVal(fArr)
        SetCellVal shp, "Prop.FirstArrivalTime", CellVal(shpFire, "Prop.FirstArrivalTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ��������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstArrivalTime" & Chr(34) & ",TheDoc!User.FirstArrivalTime)+DEPENDSON(TheDoc!User.FirstArrivalTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstArrivalTime" & Chr(34) & ", Prop.FirstArrivalTime) + DEPENDSON(Prop.FirstArrivalTime)"

    '����� ������ ������� ������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstStvolTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ������ ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ������ ������� ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FirstStvolTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 24
'        SetCellVal shp, "Prop.FirstStvolTime", GetVal(fArr)
        SetCellVal shp, "Prop.FirstStvolTime", CellVal(shpFire, "Prop.FirstStvolTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ������ ������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstStvolTime" & Chr(34) & ",TheDoc!User.FirstStvolTime)+DEPENDSON(TheDoc!User.FirstStvolTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstStvolTime" & Chr(34) & ", Prop.FirstStvolTime) + DEPENDSON(Prop.FirstStvolTime)"

    '����� �����������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LocalizationTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� �����������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� �����������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LocalizationTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 25
'        SetCellVal shp, "Prop.LocalizationTime", GetVal(fArr)
        SetCellVal shp, "Prop.LocalizationTime", CellVal(shpFire, "Prop.LocalizationTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� �����������" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LocalizationTime" & Chr(34) & ",TheDoc!User.LocalizationTime)+DEPENDSON(TheDoc!User.LocalizationTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LocalizationTime" & Chr(34) & ", Prop.LocalizationTime) + DEPENDSON(Prop.LocalizationTime)"
        
    '����� ���
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LOGTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ���������� ��" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ���������� ��������� �������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LOGTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 26
'        SetCellVal shp, "Prop.LOGTime", GetVal(fArr)
        SetCellVal shp, "Prop.LOGTime", CellVal(shpFire, "Prop.LOGTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ���������� ��" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LOGTime" & Chr(34) & ",TheDoc!User.LOGTime)+DEPENDSON(TheDoc!User.LOGTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LOGTime" & Chr(34) & ", Prop.LOGTime) + DEPENDSON(Prop.LOGTime)"
        
    '����� ���
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LPPTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ���������� ��" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ���������� ����������� ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LPPTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 27
'        SetCellVal shp, "Prop.LPPTime", GetVal(fArr)
        SetCellVal shp, "Prop.LPPTime", CellVal(shpFire, "Prop.LPPTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ���������� ��" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LPPTime" & Chr(34) & ",TheDoc!User.LPPTime)+DEPENDSON(TheDoc!User.LPPTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LPPTime" & Chr(34) & ", Prop.LPPTime) + DEPENDSON(Prop.LPPTime)"
        
    '����� ��������� ������
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireEndTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "����� ���������� �����" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "������ ��������������� ����� ���������� ����������� ������" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FireEndTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 28
'        SetCellVal shp, "Prop.FireEndTime", GetVal(fArr)
        SetCellVal shp, "Prop.FireEndTime", CellVal(shpFire, "Prop.FireEndTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "����� ���������� �����" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireEndTime" & Chr(34) & ",TheDoc!User.FireEndTime)+DEPENDSON(TheDoc!User.FireEndTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireEndTime" & Chr(34) & ", Prop.FireEndTime) + DEPENDSON(Prop.FireEndTime)"
        
End Sub


