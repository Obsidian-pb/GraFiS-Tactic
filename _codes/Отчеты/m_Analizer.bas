Attribute VB_Name = "m_Analizer"
Option Explicit
'--------------------------------------------------������ ��� ������ � �������� �� ��������----------------------------


Public Sub sP_ChangeValue(ShpObj As Visio.Shape)
'��������� ������� �� �������� ������������
Dim targetPage As Visio.Page

'---���������� ������������ ������� �������� ��� �������
    SeetsSelectForm.Show
    Set targetPage = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet)

'---��������� ��������� �������� �������
    A.Refresh targetPage.Index

'---��������� ����� ��������� ����� ������
    sP_ChangeValueMain ShpObj, targetPage

End Sub

Public Sub sP_ChangeValueMain(ByRef vsO_BaseShape As Visio.Shape, ByRef vsO_TargetPage As Visio.Page)
'��������� ������� �� �������� ������������
Dim vsO_Shape As Visio.Shape

    On Error GoTo EX

'---���������� ��� ������ � ������
    If vsO_BaseShape.Shapes.Count > 0 Then
        For Each vsO_Shape In vsO_BaseShape.Shapes
            If vsO_Shape.CellExists("Actions.ChangeValue", 0) = True Then '---��������� �������� �� ������ �������!!!
                sP_ChangeValueMain vsO_Shape, vsO_TargetPage
            End If
            If vsO_Shape.CellExists("User.PropertyValue", 0) = True Then '---��������� �������� �� ������ ����� ������
                vsO_Shape.Cells("User.PropertyValue").FormulaU = _
                    str(A.ResultByCN(vsO_Shape.Cells("Prop.PropertyName").ResultStr(visUnitsString)))
            End If
        Next vsO_Shape
    End If

Exit Sub
EX:
    SaveLog Err, "sP_ChangeValueMain"
End Sub


'------------------------��� ���������� ���������-----------------------------------------
Public Sub sP_ChangeRDValue(ShpObj As Visio.Shape)
'��������� ������� ���������� ������ � ���������� ���������
Dim targetPage As Visio.Page
'Dim tmp As Variant

'---���������� ������������ ������� �������� ��� �������
    SeetsSelectForm.Show
    Set targetPage = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet)

'---��������� ��������� �������� �������
    A.Refresh targetPage.Index

'---��������� ������������� ������ ��� ����� ���������� ���������
    '---������� �������
    ShpObj.Cells("Prop.Personnel").Formula = """" & A.Result("PersonnelHave") & "/" & A.Result("PersonnelNeed") & """"
    '---�������� ��
    ShpObj.Cells("Prop.MainPA").Formula = """" & A.Result("MainPAHave") & "/" & A.Result("ACNeed") & """"
    '---������ ����
    ShpObj.Cells("Prop.WaterExpense").Formula = """" & A.Result("FactStreamW") & "/" & A.Result("NeedStreamW") & """"
    '---����� ���� (���� ���� ����������� ���������, �� �������� ���������� ������ ���� ������ ����������� ������ ������������)
'    tmp = A.Result("WaterValueNeed10min")
    If A.Result("WaterEternal") Then
        ShpObj.Cells("Prop.WaterValue").Formula = """" & A.Result("WaterValueHave") & "/" & A.Result("WaterValueHave") & """"
    Else
        ShpObj.Cells("Prop.WaterValue").Formula = """" & A.Result("WaterValueHave") & "/" & A.Result("WaterValueNeed10min") & """"
    End If
    '---������� ����
    ShpObj.Cells("Prop.GDZS").Formula = """" & A.Sum("GDZSChainsCountWork;GDZSChainsRezCountHave") & "/" & A.Result("GDZSChainsCountNeed") & """"
    '---�������
    ShpObj.Cells("Prop.Stv").Formula = """" & A.Result("StvolWHave") & "/" & PF_RoundUp(A.Result("NeedStreamW") / ShpObj.Cells("Prop.StvWaterExpense").Result(visNumber)) & """"

End Sub
