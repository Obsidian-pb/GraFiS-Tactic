VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_ManagementGDZS 
   Caption         =   "UserForm1"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13170
   OleObjectBlob   =   "f_ManagementGDZS.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_ManagementGDZS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1



Private Sub ListView1_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    
    Application.ActiveWindow.Select Application.ActivePage.Shapes(ListView1.SelectedItem.Key), visDeselectAll
    Application.ActiveWindow.Select Application.ActivePage.Shapes(ListView1.SelectedItem.Key), visSelect
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ps_SortBy ColumnHeader.Index + 1
End Sub

Private Sub ListView1_DblClick()
    Application.ActiveWindow.Zoom = 1.5 * pf_GetScaleAt200
    Application.ActiveWindow.ScrollViewTo Application.ActivePage.Shapes(ListView1.SelectedItem.Key).Cells("PinX"), Application.ActivePage.Shapes(ListView1.SelectedItem.Key).Cells("PinY")
End Sub

Public Sub pS_Stretch()
    Me.ListView1.Width = Me.InsideWidth - 6
    Me.ListView1.Height = Me.InsideHeight - 6
End Sub

Private Sub UserForm_Deactivate()
    MngmnGDZSWndwHide
End Sub

Private Sub UserForm_Resize()
    pS_Stretch
End Sub



'-------------------------------------����������------------------------------------
Private Sub ps_SortBy(ByVal ColumnNumber As Integer)
'��������� ���������� ������ � ListView �� ���������� �������
Dim DataArray() As String
Dim TempString() As String '��������� ������ ��� ����������
Dim i As Integer
Dim j As Integer
Dim k As Integer

On Error GoTo EX

ReDim DataArray(Me.ListView1.ListItems.Count, Me.ListView1.ListItems(1).ListSubItems.Count + 2) '2 - ������ ��� 1 ��������� ��� ������������� ��������, 2 - ��� �������� Key
ReDim TempString(Me.ListView1.ListItems(1).ListSubItems.Count + 2)

For i = 1 To Me.ListView1.ListItems.Count '���������� ������
    DataArray(i, 1) = Me.ListView1.ListItems(i).Key
    DataArray(i, 2) = Me.ListView1.ListItems(i)
    For j = 3 To Me.ListView1.ListItems(1).ListSubItems.Count + 2 '���������� �������
        DataArray(i, j) = Me.ListView1.ListItems(i).ListSubItems(j - 2)
    Next j
Next i

'---��������� ����������
If ColumnNumber = 2 Or ColumnNumber = 3 Or ColumnNumber = 4 Or ColumnNumber = 7 Then
    '��� �����
    For i = 1 To Me.ListView1.ListItems.Count '������ �����
        For j = i To Me.ListView1.ListItems.Count '������ �����
            If DataArray(j, ColumnNumber) < DataArray(i, ColumnNumber) Then
                '��������� ������ ������ �� ��������� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    TempString(k) = DataArray(i, k)
                Next k
                '��������� ������ ������ � ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(i, k) = DataArray(j, k)
                Next k
                '��������� �������� ������� �� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(j, k) = TempString(k)
                Next k
            End If
        Next j
    Next i
ElseIf ColumnNumber = 5 Then
    '��� ���
    For i = 1 To Me.ListView1.ListItems.Count '������ �����
        For j = i To Me.ListView1.ListItems.Count '������ �����
            If CDate(DataArray(j, ColumnNumber)) < CDate(DataArray(i, ColumnNumber)) Then
                '��������� ������ ������ �� ��������� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    TempString(k) = DataArray(i, k)
                Next k
                '��������� ������ ������ � ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(i, k) = DataArray(j, k)
                Next k
                '��������� �������� ������� �� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(j, k) = TempString(k)
                Next k
            End If
        Next j
    Next i
ElseIf ColumnNumber = 6 Then
    '��� �����
    For i = 1 To Me.ListView1.ListItems.Count '������ �����
        For j = i To Me.ListView1.ListItems.Count '������ �����
            If CDec(DataArray(j, ColumnNumber)) < CDec(DataArray(i, ColumnNumber)) Then
                '��������� ������ ������ �� ��������� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    TempString(k) = DataArray(i, k)
                Next k
                '��������� ������ ������ � ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(i, k) = DataArray(j, k)
                Next k
                '��������� �������� ������� �� ������
                For k = 1 To Me.ListView1.ListItems(1).ListSubItems.Count + 2
                    DataArray(j, k) = TempString(k)
                Next k
            End If
        Next j
    Next i
End If

'---��������� �������
Me.ListView1.ListItems.Clear
For i = 1 To UBound(DataArray, 1) '������ ����� - ��������� ������
    Me.ListView1.ListItems.Add i, DataArray(i, 1), DataArray(i, 2)
    For j = 3 To UBound(DataArray, 2) '������ ����� - ���������� ������
        Me.ListView1.ListItems(i).ListSubItems.Add j - 2, , DataArray(i, j)
    Next j
Next i

Exit Sub
EX:
'������
End Sub


'-------------------------------------�������---------------------------------------
Private Function pf_GetScaleAt200() As Double
'���������� ����������� ���������� ������� ������� �������� ������������ �������� 1:200
Dim v_Minor As Double
Dim v_Major As Double

    v_Minor = Application.ActivePage.PageSheet.Cells("PageScale").Result(visNumber)
    v_Major = Application.ActivePage.PageSheet.Cells("DrawingScale").Result(visNumber)
    pf_GetScaleAt200 = (v_Major / v_Minor) / 200
End Function


