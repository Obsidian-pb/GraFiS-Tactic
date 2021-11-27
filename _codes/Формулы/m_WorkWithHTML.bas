Attribute VB_Name = "m_WorkWithHTML"
Option Explicit



Public Sub ShowData(ByRef shp As Visio.Shape, ByVal htmlText As String)
    f_FormulaForm.ShowHTML shp, htmlText
End Sub

Public Sub ShowAllData(ByVal htmlText As String)
    f_FormulaFormAll.ShowData htmlText
End Sub

Public Sub ShowDataInShape(ByRef shp As Visio.Shape, ByVal htmlText As String)
    On Error Resume Next
    f_FormulaForm.CopyBrowserContent shp, htmlText
'    SetShowDataInShapeControlFormula shp, "User.DataChangeAction.Prompt"
End Sub

Private Sub SetShowDataInShapeControlFormula(ByRef shp As Visio.Shape, ByVal cellName As String)
Dim i As Integer
Dim frml As String
Dim nameOfRow As String
    
    frml = ""
    For i = 0 To shp.RowCount(visSectionProp) - 1
        '�������� ��� ������
        nameOfRow = shp.CellsSRC(visSectionProp, i, visCustPropsValue).RowNameU
             
        '���������� �������� ������� ��� �������� ��������� (��� ���������� � ������ ������)
        If Len(frml) > 0 Then frml = frml & "&"
        frml = frml & "Prop." & nameOfRow
    Next i
    
    SetCellFrml shp, cellName, frml
End Sub

Public Sub ClearShapeText(ByRef shp As Visio.Shape)
    shp.Text = " "
End Sub

Public Sub TryGetFromAnalaizer(ByRef shp As Visio.Shape, ByVal nameOfRow As String)
    If left(nameOfRow, 2) = "A_" Then
        SetCellVal shp, "Prop." & nameOfRow, a.Result(Right(nameOfRow, Len(nameOfRow) - 2))
    End If
End Sub




Public Sub ShowAElemAddForm(ByRef shp As Visio.Shape)
    f_InsertAnalizedCell.ShowMe shp
End Sub

Public Sub ShowLinkToCellForm(ByRef shp As Visio.Shape)
    f_LinkToCell.ShowMe shp
End Sub


Public Function PatternToHTML(ByRef shp As Visio.Shape, ByVal htmlText As String) As String
'������� ������ ��� ��������� HTML � �������� ����������� ������ �� ������ ������ shp ������������ �� ����������
Dim cll As Visio.cell
Dim i As Integer
Dim targetPageIndex As Integer
    
    '��������� ���������� "�������"
    targetPageIndex = cellVal(shp, "Prop.TargetPageIndex")
    If targetPageIndex = 0 Then
        a.Refresh Application.ActivePage.Index
    Else
        a.Refresh targetPageIndex
    End If
    
    
    htmlText = Replace(htmlText, Asc(34), "'")
    
    For i = 0 To shp.RowCount(visSectionProp) - 1
        '�������� ������ �� ������ ������ Props
        Set cll = shp.CellsSRC(visSectionProp, i, visCustPropsValue)
        '�������� �������� ������ �� ����������� "�������"
        TryGetFromAnalaizer shp, cll.RowNameU
        '��������� �������� ����� html
        htmlText = Replace(htmlText, "$" & cll.RowNameU & "$", ClearString(cll.ResultStr(visUnitsString)))
    Next i

PatternToHTML = htmlText
End Function



'--------------���������� ���� ������----------------
Public Sub RefreshAllFormulas()
Dim frmlCollection As Collection
Dim shp As Visio.Shape
    
    Set frmlCollection = GetGFSShapes(Array("User.IndexPers:500"))
    For Each shp In frmlCollection
        ShowDataInShape shp, cellVal(shp, "User.TextPattern", visUnitsString)
    Next shp
End Sub

'--------------����� ������ � ����� ����--------
Public Sub ShowAllFormulas()
Dim frmlCollection As Collection
Dim shp As Visio.Shape
Dim htmlText As String
    
    Set frmlCollection = GetGFSShapes(Array("User.IndexPers:500"))
    
    '��������� ������ � ��������� �� ������ �� ���������
    Set frmlCollection = Sort(frmlCollection, "PinY")
    
    '��������� �������� ������ html ����
    For Each shp In frmlCollection
        htmlText = htmlText + PatternToHTML(shp, cellVal(shp, "User.TextPattern", visUnitsString))
'        ShowDataInShape shp, cellVal(shp, "User.TextPattern", visUnitsString)
    Next shp
    
    ShowAllData htmlText
End Sub

'-------------------------------���������� �������� �����----------------------------------------------
Public Function Sort(ByVal shps As Collection, ByVal sortCellName As String) As Collection
'������� ���������� ��������������� ��������� �����. ������ ����������� �� �������� � ������ sortCellName - ��� ������, ��� ����
Dim i As Integer
Dim tmpshp As Visio.Shape
Dim tmpColl As Collection

    
    Set tmpColl = New Collection
    
    Do While shps.Count > 1
        
        Set tmpshp = GetMaxShp(shps, sortCellName)
        
        AddUniqueCollectionItem tmpColl, tmpshp
        RemoveFromCollection shps, tmpshp
        
        i = i + 1
        If i > 100 Then Exit Do
    Loop
    
    Set tmpshp = shps(1)
    AddUniqueCollectionItem tmpColl, tmpshp
    Debug.Print tmpshp.ID & " " & cellVal(tmpshp, sortCellName)
    
    Set Sort = tmpColl
End Function

Public Function GetMaxShp(ByRef col As Collection, ByVal sortCellName As String) As Visio.Shape
Dim i As Integer
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim shp1Val As Single
Dim shp2Val As Single
    
    Set shp1 = col(1)
    shp1Val = cellVal(shp1, sortCellName)
    For i = 1 To col.Count
        Set shp2 = col(i)
        shp2Val = cellVal(shp2, sortCellName)
        
        If shp2Val > shp1Val Then
            Set shp1 = shp2
            shp1Val = shp2Val
        End If
    Next i
    Debug.Print shp1.ID & " " & shp1Val
    
Set GetMaxShp = shp1
End Function
















