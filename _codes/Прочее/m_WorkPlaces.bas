Attribute VB_Name = "m_WorkPlaces"
Option Explicit




Public Sub PS_AddWorkPlaces(ShpObj As Visio.Shape)
'��������� ��������� ������ ������� ���� � ������������ � ���������
Dim WorkPlaceBuilder As c_WorkPlaces

'---��������� ������� ��������� WALL_M.VSS
     If PF_DocumentOpened("WALL_M.VSS") = False Then ' And PF_DocumentOpened("WALL_M.VSSX") = False Then
        ShpObj.Delete
        MsgBox "�������� '����������� ��������' �� ���������! ���������� ������� ����������!'", vbCritical
        Exit Sub
     End If
    
'---���������� ������
    Set WorkPlaceBuilder = New c_WorkPlaces
        WorkPlaceBuilder.S_SetFullShape
    Set WorkPlaceBuilder = Nothing
    
'---������� ��������� ������
    ShpObj.Delete

End Sub

Public Sub PS_WorkPlacesRenum(ShpObj As Visio.Shape)
'��������� ���������� ��� ��������� �� ����� ������ ������� ���� � ���������������� ��
Dim WorkPlace As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim i As Integer

'---�������� ��� ������ ���� "������� �����"
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "�����")
    Application.ActiveWindow.Selection = vsoSelection
    
'---���������� � ���������������� ������ ������
    i = 1
    For Each WorkPlace In vsoSelection
        WorkPlace.Cells("Prop.LocationID").FormulaU = i
        i = i + 1
    Next
    
'---������� ��������� ������
    ShpObj.Delete

End Sub

'-----------------------------------������� �����������-----------------------------------------
Public Sub PS_AddExplicationTable(ShpObj As Visio.Shape)
'����� ��������� ������� �����������
Dim vO_Shape As Visio.Shape
Dim vO_NewString As Visio.Shape
Dim vO_StringMaster As Visio.Master
Dim i As Integer

Dim colWorkplaces As Collection

On Error GoTo Tail
    
    '---������� ��������� ����
    Set colWorkplaces = New Collection
    '---��������� ��������� ����
    For Each vO_Shape In Application.ActivePage.Shapes
        If PFB_isPlace(vO_Shape) Then
            colWorkplaces.Add vO_Shape
        End If
    Next vO_Shape
    
    '---���� � ��������� ��� ����� - �������
        If colWorkplaces.Count = 0 Then Exit Sub
    
    '---��������� ���������
    sC_SortPlaces colWorkplaces
    
    '---��������� ���� ������� ����������� Visio
    Application.EventsEnabled = False
    
    '---�������� ������ �����
    Set vO_StringMaster = ThisDocument.Masters("�����������")
    
    '---���������� ��� ������ �� �����, � ���� ������ - ������ �����, ������� ������ �����������
    i = 0
    For Each vO_Shape In colWorkplaces
        If PFB_isPlace(vO_Shape) Then
            i = i + 1
            
            If i = 1 Then
                '---����������� �������� ��������� ������
                ShpObj.Cells("User.PlaceSheetName").FormulaU = """" & vO_Shape.NameID & """"
            Else
                '---��������� ����� �������
                Set vO_NewString = Application.ActivePage.Drop(vO_StringMaster, _
                                    ShpObj.Cells("PinX").Result(visInches), _
                                    ShpObj.Cells("PinY").Result(visInches) - ShpObj.Cells("Height").Result(visInches) * (i - 1))
                vO_NewString.Cells("User.PlaceSheetName").FormulaU = """" & vO_Shape.NameID & """"
            End If
        End If
    Next vO_Shape

    '---�������������� ����
    Application.EventsEnabled = True
    Set colWorkplaces = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_AddExplicationTable"
    Application.EventsEnabled = True
    Set colWorkplaces = Nothing
End Sub


Private Sub sC_SortPlaces(ByRef aC_PlacesCollection As Collection)
'��������� ���������� ��������� � ��������� �������� ������ �� ����������� ������� �������
Dim vsi_MinKod, vsi_CurKod As Integer
'Dim vO_Place1 As Visio.Shape
'Dim vO_Place2 As Visio.Shape
Dim vsCol_TempCol As Collection
Dim i, k, j As Integer

'---��������� ����� ���������
Set vsCol_TempCol = New Collection

'---��������� ���� ���������� �� ����� ���������� ������� ������������ ���������� ��-��� � ���������
For i = 1 To aC_PlacesCollection.Count
    vsi_MinKod = aC_PlacesCollection.Item(1).Cells("Prop.LocationID").Result(visNumber)
    j = 1
    '---��������� ���� ���������� �� ����� ���������� ������� �������� ���������� ��-��� � ���������
    For k = 1 To aC_PlacesCollection.Count
        vsi_CurKod = aC_PlacesCollection.Item(k).Cells("Prop.LocationID").Result(visNumber)
        If vsi_CurKod < vsi_MinKod Then '���� ������� ��� ������ ������������, �� ���������� ����� �������
            vsi_MinKod = vsi_CurKod
            j = k
        End If
    Next k
    '---��������� �� ��������� ������ ��-� � ���������� ��������� �������
    vsCol_TempCol.Add aC_PlacesCollection.Item(j), str(aC_PlacesCollection.Item(j).ID)
    aC_PlacesCollection.Remove j '�� �������� ��������� - ������� ���
    
Next i


'---��������� ����������� ��������� � ������������ � ���������� ���������
Set aC_PlacesCollection = vsCol_TempCol

Set vsCol_TempCol = Nothing
End Sub





