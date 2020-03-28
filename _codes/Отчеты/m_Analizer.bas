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

