Attribute VB_Name = "m_ExportDescriptions"
Option Explicit

'Private exl As Object ' Excel.Application
'Private wkbk As Object '  Excel.Workbook
'Private wkst As Object '  Excel.Worksheet
Private cmndID As Integer




'Public Sub DescriptionExportToWord()
''������������ �������� ������ �������� � �������� Word
'Dim wrd As Object
'Dim wrdDoc As Object
'Dim wrdTbl As Object
'Dim wrdTblRow As Object
''Dim comRowsCount As Integer
'Dim i As Integer
'Dim r As Integer
'Dim shp As Visio.Shape
'Dim comCol As Collection
'Dim comColSorted As Collection
'Dim com As c_SimpleDescription
'Dim curTime As Date
'Dim fireTime As Date
'
'    '---��������� ��������� ����� � ��������� �� �� �������
'    Set comCol = New Collection
'    cmndID = 0
'    For Each shp In Application.ActivePage.Shapes
'        checkCommands comCol, shp
'    Next shp
'
'    '---���������
'    Set comColSorted = SortCommands(comCol)
'
'
'    '������� ����� �������� Word
'    Set wrd = CreateObject("Word.Application")
'    wrd.Visible = True
'    wrd.Activate
'    Set wrdDoc = wrd.Documents.Add
'    wrdDoc.Activate
'    '������� � ����� ��������� ������� ��������� ��������
''    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, comColSorted.Count, 9)
'    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, 1, 9)
'    With wrdTbl
'        If .style <> "����� �������" Then
'            .style = "����� �������"
'        End If
'        .ApplyStyleHeadingRows = True
'        .ApplyStyleLastRow = False
'        .ApplyStyleFirstColumn = True
'        .ApplyStyleLastColumn = False
'        .ApplyStyleRowBands = True
'        .ApplyStyleColumnBands = False
'    End With
'
'
'    '��������� ������� �������� ������ ��������
'    If comColSorted.Count > 0 Then
'        r = 1
'        A.Refresh Application.ActivePage.Index, curTime
''        fireTime = A.Result("FireTime")     '�� ������� �������
'        fireTime = cellval(Application.ActivePage.Shapes, "Prop.FireTime", visUnitsString)  ' visDate)    '��� ��������
'        curTime = comColSorted(1).time
'        '---������� ������ ������
'        If DateDiff("n", fireTime, curTime) < 2000 Then
'            wrdTbl.Rows(r).Cells(1).Range.text = "�+" & DateDiff("n", fireTime, curTime)
'        Else
'            wrdTbl.Rows(r).Cells(1).Range.text = Format(curTime, "HH:MM")
'        End If
'        wrdTbl.Rows(r).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
'        wrdTbl.Rows(r).Cells(4).Range.text = A.Result("StvolWBHave")
'        wrdTbl.Rows(r).Cells(5).Range.text = A.Result("StvolWAHave")
'        wrdTbl.Rows(r).Cells(6).Range.text = A.Result("StvolWLHave")
'        wrdTbl.Rows(r).Cells(7).Range.text = A.Result("StvolFoamHave")
'        wrdTbl.Rows(r).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
'        Set com = comColSorted(1)
'        If com.sdType = 1 Then
'            wrdTbl.Rows(r).Cells(9).Range.text = com.text
'        Else
'            wrdTbl.Rows(r).Cells(2).Range.text = com.text
'        End If
'
'        For i = 1 To comColSorted.Count
'            Set com = comColSorted(i)
'            '���� ����� ������� ������ ������ ��� ����� ���������� - ��������� ����������
'            If curTime <> com.time Then
'                A.Refresh Application.ActivePage.Index, com.time
'                '---������� ������� ������� �� ������ (��� ������� �� ���������)
'                r = wrdTbl.Rows.Add().Index
'
'                wrdTbl.Rows(r).Cells(1).Range.text = "�+" & DateDiff("n", fireTime, com.time)
'                wrdTbl.Rows(r).Cells(3).Range.text = Round(A.Result("NeedStreamW"), 1)
'                wrdTbl.Rows(r).Cells(4).Range.text = A.Result("StvolWBHave")
'                wrdTbl.Rows(r).Cells(5).Range.text = A.Result("StvolWAHave")
'                wrdTbl.Rows(r).Cells(6).Range.text = A.Result("StvolWLHave")
'                wrdTbl.Rows(r).Cells(7).Range.text = A.Result("StvolFoamHave")
'                wrdTbl.Rows(r).Cells(8).Range.text = Round(A.Result("FactStreamW"), 1)
'
'                curTime = comColSorted(i).time
'            End If
'
'            If com.sdType = 1 Then
'                If Asc(wrdTbl.Rows(r).Cells(9).Range.text) = 13 Then
'                    wrdTbl.Rows(r).Cells(9).Range.text = com.text
'                Else
'                    wrdTbl.Rows(r).Cells(9).Range.text = wrdTbl.Rows(r).Cells(9).Range.text & Chr(10) & com.text
'                End If
'            Else
'                If Asc(wrdTbl.Rows(r).Cells(9).Range.text) = 13 Or Asc(wrdTbl.Rows(r).Cells(9).Range.text) = 10 Then
'                    wrdTbl.Rows(r).Cells(2).Range.text = com.text
'                Else
'                    wrdTbl.Rows(r).Cells(2).Range.text = wrdTbl.Rows(r).Cells(2).Range.text & Chr(10) & com.text
'                End If
'            End If
'
'        Next i
'    End If
'    wrdTbl.AutoFitBehavior 1        '������������� ������ �������� �� �����������
'End Sub

Public Sub DescriptionExportToWord()
'������������ �������� ������ �������� � �������� Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
Dim wrdTblRow As Object
'Dim comRowsCount As Integer
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim shp As Visio.Shape
Dim comCol As Collection
Dim comColSorted As Collection
Dim com As c_SimpleDescription
Dim curTime As Date
Dim fireTime As Date

Dim strArr(9) As String
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set comCol = New Collection
    cmndID = 0
    For Each shp In Application.ActivePage.Shapes
        checkCommands comCol, shp
    Next shp
    
    '---���������
    Set comColSorted = SortCommands(comCol)
    
    
    '������� ����� �������� Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add
    wrdDoc.Activate
    '������� � ����� ��������� ������� ��������� ��������
'    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, comColSorted.Count, 9)
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, 1, 9)
    With wrdTbl
        If .style <> "����� �������" Then
            .style = "����� �������"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    
    
    '��������� ������� �������� ������ ��������
    If comColSorted.Count > 0 Then
        r = 1
        a.Refresh Application.ActivePage.Index, curTime
'        fireTime = A.Result("FireTime")     '�� ������� �������
        fireTime = cellVal(Application.ActivePage.Shapes, "Prop.FireTime", visUnitsString)  ' visDate)    '��� ��������
        curTime = comColSorted(1).time
'        '---������� ������ ������
'        htmlTable = htmlTable & "<thead><tr>"
'            htmlTable = htmlTable & "<th>1</th>"
'            htmlTable = htmlTable & "<th>2</th>"
'            htmlTable = htmlTable & "<th>3</th>"
'            htmlTable = htmlTable & "<th>4</th>"
'            htmlTable = htmlTable & "<th>5</th>"
'            htmlTable = htmlTable & "<th>6</th>"
'            htmlTable = htmlTable & "<th>7</th>"
'            htmlTable = htmlTable & "<th>8</th>"
'            htmlTable = htmlTable & "<th>9</th>"
'        htmlTable = htmlTable & "</tr></thead>" & Chr(10)
        
        '������ ��� ������ ������
        Set com = comColSorted(1)
        a.Refresh Application.ActivePage.Index, com.time
        strArr(1) = "�+" & DateDiff("n", fireTime, com.time)
        strArr(3) = Round(a.Result("NeedStreamW"), 1)
        strArr(4) = a.Result("StvolWBHave")
        strArr(5) = a.Result("StvolWAHave")
        strArr(6) = a.Result("StvolWLHave")
        strArr(7) = a.Result("StvolFoamHave")
        strArr(8) = Round(a.Result("FactStreamW"), 1)
        If com.sdType = 1 Then
            strArr(9) = strArr(9) & com.text & Chr(10) & Chr(10)
        Else
            strArr(2) = strArr(2) & com.text & Chr(10) & Chr(10)
        End If
        curTime = com.time

        '��� ����������� ������
        For i = 2 To comColSorted.Count
            Set com = comColSorted(i)
            If curTime = com.time Then
                If com.sdType = 1 Then
                    strArr(9) = strArr(9) & com.text & Chr(10) & Chr(10)
                Else
                    strArr(2) = strArr(2) & com.text & Chr(10) & Chr(10)
                End If
            Else
                '��������� ���!
                For j = 1 To UBound(strArr)
                    wrdTbl.Rows(r).Cells(j).Range.text = strArr(j)
                Next j
                
                '��������� ������ ��� �������� �������
                ClearArr strArr
                
                a.Refresh Application.ActivePage.Index, com.time
                
                strArr(1) = "�+" & DateDiff("n", fireTime, com.time)
                strArr(3) = Round(a.Result("NeedStreamW"), 1)
                strArr(4) = a.Result("StvolWBHave")
                strArr(5) = a.Result("StvolWAHave")
                strArr(6) = a.Result("StvolWLHave")
                strArr(7) = a.Result("StvolFoamHave")
                strArr(8) = Round(a.Result("FactStreamW"), 1)
                
                If com.sdType = 1 Then
                    strArr(9) = strArr(9) & com.text & Chr(10) & Chr(10)
                Else
                    strArr(2) = strArr(2) & com.text & Chr(10) & Chr(10)
                End If
                
                curTime = com.time
                r = wrdTbl.Rows.Add().Index
            End If
        Next i

    End If
    
    wrdTbl.AutoFitBehavior 1        '������������� ������ �������� �� �����������
End Sub

Private Function ClearArr(ByRef arr As Variant) As Variant
Dim i As Integer
    
    For i = 0 To UBound(arr)
        arr(i) = ""
    Next i
    
ClearArr = arr
End Function


Public Sub DescriptionViewInList()
'������������ �������� ������ �������� � �������� Word
Dim i As Integer
Dim shp As Visio.Shape
Dim comCol As Collection
Dim comColSorted As Collection
Dim com As c_SimpleDescription
Dim curTime As Date
Dim fireTime As Date

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---��������� ��������� ����� � ��������� �� �� �������
    Set comCol = New Collection
    cmndID = 0
    For Each shp In Application.ActivePage.Shapes
        checkCommands comCol, shp
    Next shp
    
    
    
    '---���������
    Set comColSorted = SortCommands(comCol)
    
  
    
    '��������� ������� �������� ������ ��������
    If comColSorted.Count > 0 Then
        a.Refresh Application.ActivePage.Index, curTime
        
        fireTime = cellVal(Application.ActivePage.Shapes, "Prop.FireTime", visUnitsString)  ' visDate)    '��� ��������
'        fireTime = A.Result("FireTime")
        curTime = comColSorted(1).time
        
        ReDim myArray(comColSorted.Count, 3)
        '---������� ������ ������
        myArray(0, 0) = "ID"
        myArray(0, 1) = "�����"
        myArray(0, 2) = "��������� ����������"
        myArray(0, 3) = "�������/��������"
    
        For i = 1 To comColSorted.Count
            Set com = comColSorted(i)
                curTime = comColSorted(i).time
                
                myArray(i, 0) = com.shp.ID
                If DateDiff("n", fireTime, curTime) < 2000 Then
                    myArray(i, 1) = "�+" & DateDiff("n", fireTime, curTime)
                Else
                    myArray(i, 1) = Format(curTime, "HH:MM")
                End If
                
                If com.sdType = 1 Then
                    myArray(i, 3) = com.text
                Else
                    myArray(i, 2) = com.text
                End If
        Next i
    End If

    '---���������� �����
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;300 pt;300 pt", "InfoAndCommands", "���������� � �������"

End Sub










Private Sub checkCommands(ByRef comCol As Collection, ByRef shp As Visio.Shape)
Dim i As Integer
Dim rowName As String
Dim cmnd As c_SimpleDescription
Dim tmpStr As String
    
    On Error GoTo ex
    
    For i = 0 To shp.RowCount(visSectionUser) - 1
        rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
        If Len(rowName) > 9 Then
            If Left(rowName, 12) = "GFS_Command_" Then
                Set cmnd = New c_SimpleDescription
                cmnd.Activate shp, shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), CStr(cmndID)
                AddUniqueCollectionItem comCol, cmnd
'                        shp.Cells("Actions." & rowName & ".Action").Formula = Replace(shp.Cells("Actions." & rowName & ".Action").Formula, "���", "����������_���")
                
                cmndID = cmndID + 1
            End If
            
            If Left(rowName, 9) = "GFS_Info_" Then
                tmpStr = shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString)
                If Not InStr(1, tmpStr, " ������ ") > 0 Then
                    Set cmnd = New c_SimpleDescription
                    cmnd.ActivateAsInfo shp, shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString), CStr(cmndID)
                    AddUniqueCollectionItem comCol, cmnd

                    cmndID = cmndID + 1
                End If
            End If
        End If
    Next i
    
    
Exit Sub
ex:

End Sub



Public Function getCallName(ByRef shp As Visio.Shape) As String
Dim txt As String

    On Error GoTo ex
    getCallName = ""
    '�������������
    txt = shp.Cells("Prop.Unit").ResultStr(visUnitsString)
    If Not txt = "" Then
        getCallName = txt
    End If
    '��������
    txt = shp.Cells("Prop.Call").ResultStr(visUnitsString)
    If Not txt = "" Then
        getCallName = getCallName & "(" & txt & ")"
    End If
    
Exit Function
ex:
    getCallName = "-"
End Function


