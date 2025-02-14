VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementsShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================����-������� ��� ��������� ���������==========================================
Private elemCollection As Collection            '������ ���������
Private elemWithCellsCollection As Collection   '��������� ��������� ��� ������� ������  ����� ��������� �����

'-------------------------------------��������---------------------------------



'-------------------------------------������� ��������� �������------------------------------------
Private Sub Class_Initialize()
    Set elemCollection = New Collection
    Set elemWithCellsCollection = New Collection
    fillElemCollection
End Sub
Private Sub Class_Terminate()
    Set elemCollection = Nothing
    Set elemWithCellsCollection = Nothing
End Sub

'-------------------------------------�������� ������----------------------------------------------
Public Property Get count() As Integer
'���������� ��������� � ���������
    count = elemCollection.count
End Property




'-------------------------------------���������� ��������� � ������� �������----------------------------------------------
Private Sub fillElemCollection()
'��������� ��������� ���������
    AddNewElement "RSCHSTechincsHave", "������� ���� �������", "{$} ({TotalFireCount} - �������-������������)"
    AddNewElement "MainPAHave", "�� ������ ���������� �������", , , , , False
    AddNewElement "TargetedPAHave", "�� �������� ���������� �������", , , , , False
    AddNewElement "MainOverallHave", "�� �������� �������", "{$} ({TargetedPAHave} - ���.����., {MainPAHave} - ���.����.)"
    AddNewElement "PAHighHave", "�� �������� �������", , , , , False
    AddNewElement "SpecialPAHave", "�� ����������� �������", "{$} ({PAHighHave} - ��������)"
    AddNewElement "OtherTechincsHave", "������ �������-������������ ������� �������"
    
    AddNewElement "ACCount", "�� �������", , , , , False
    AddNewElement "ACNeed", "�� ���������", , , , , False
    AddNewElement "ANRNeed", "��� ���������", , , , , False
    AddNewElement "AGCount", "����� �������", , , , , False
    AddNewElement "ALCount", "�� �������", , , , , False
    AddNewElement "AKPCount", "��� �������", , , , , False
    AddNewElement "MVDCount", "������� ��� �������", , , , , False
    AddNewElement "SMPCount", "������� ��� �������", , , , , False
    AddNewElement "TechnicsNotMchsHave", "������� ������ ����� �������", "{$} ({MVDCount} - ���, {SMPCount} - ���)"
    AddNewElement "TechnicsNotMchsSmpMvdHave", "������� ������� ����� ����, ���, ���"
    AddNewElement "TechnicsNotMchsOtherHave", "������ ������� �������"
    AddNewElement "TotalTechCount", "������� ����� �������", , , , , False
    AddNewElement "TotalFireCount", "�������� � ��������-������������ ������� ����� �������", , , , , False
    
    AddNewElement "FireSquare", "������� ������", "{$} �.��. ({ExtSquare} �.��. - ������� �������)", , "User.FireSquare"
    AddNewElement "ExtSquare", "������� �������", , , "User.ExtSquare", , False
    AddNewElement "FactStreamW", "������ ���� �����, �����������", "{$} �/�", , "User.PodOut"
    AddNewElement "FactStreamWExt", "������ ���� �� �������, �����������", , , , , False
    AddNewElement "FactStreamWProt", "������ ���� �� ������, �����������", , , , , False
    AddNewElement "FactStreamWChill", "������ ���� �� ����������, �����������", , , , , False
    AddNewElement "IntenseWater", "������������� ������ ����"
    AddNewElement "NeedStreamW", "������ ���� �����, ���������", "{$} �/�", , "User.NeedStreamW", , True
    
    AddNewElement "PersonnelHave", "������� ������� �������", "{$} ���. ({PersonnelNeed} ���. - ���������)"
    AddNewElement "PersonnelNeed", "������� ������� ���������", , , "Prop.Personnel", , False
    
    AddNewElement "StvolWHave", "������� ������� ������", "{$} ({StvolWBHave} - ���. '�', {StvolWAHave} - ���. '�', {StvolWLHave} - ���. ���.)"
    AddNewElement "StvolWBHave", "������� � ������", , , , , False
    AddNewElement "StvolWAHave", "������� � ������", , , , , False
    AddNewElement "StvolWLHave", "������� ������� �������� ������", , , , , False
    AddNewElement "StvolWBNeed", "������� � ���������", , , , , False
    AddNewElement "StvolWANeed", "������� � ���������", , , , , False
    AddNewElement "StvolWLNeed", "������� ������� �������� ���������", , , , , False
    AddNewElement "StvolHave", "������� ����� ������", "{$} ({StvolExtHave} - �������, {StvolProtHave} - ������, {StvolChillHave} - ����������, {StvolEvacHave} - ���������, {StvolCheckHave} - ��������)", , , , True
    AddNewElement "StvolExtHave", "������� �� �������, ������", , , , , False
    AddNewElement "StvolProtHave", "������� �� ������, ������", , , , , False
    AddNewElement "StvolChillHave", "������� �� ����������, ������", , , , , False
    AddNewElement "StvolEvacHave", "������� ��� ���������, ������", , , , , False
    AddNewElement "StvolCheckHave", "������� �� ��������, ������", , , , , False
    
    AddNewElement "StvolFoamHave", "������� ������ ������"
    AddNewElement "StvolGasHave", "������� ������� ������"
    AddNewElement "StvolPowderHave", "������� ���������� ������"
    AddNewElement "Razvetvlen", "������������", , , , , False

    AddNewElement "GetedWaterValue", "���������� ���� �� ������� ����������", "{$} �/� (max = {GetedWaterValueMax} �/� )", , , , False
    AddNewElement "GetedWaterValueMax", "���������� ���� �� ������� ���������� ��������", , , , , False
    AddNewElement "GetingWaterCount", "��� ����������� �� �������������", "{$} ({PANeedOnWaterSource} - ��������� ����������)"
    AddNewElement "PANeedOnWaterSource", "��������� ���������� �� �������������", , , , , False
    AddNewElement "WaterEternal", "����������� ����� ����", , False, , False, False

    AddNewElement "GDZSChainsCountWork", "������� ���� ��������", "{$} ({GDZSMansCountWork} ������������������)"
    AddNewElement "GDZSChainsCountWorkSoft", "������� ���� �������� � ����������� ��������", , , , , False
    AddNewElement "GDZSChainsCountWorkHard", "������� ���� �������� � ������� ��������", , , , , False
    AddNewElement "GDZSChainsCountNeed", "������� ���� ��������� (+ ������)"
    AddNewElement "GDZSChainsRezCountHave", "������� ���� � ������� �������", "{$} ({GDZSMansRezCountHave} ������������������)", , , , False
    AddNewElement "GDZSChainsRezCountNeed", "������� ���� � ������� ���������", , , , , False
'    AddNewElement "GDZSChainsCanBeFormed", "������� ���� ����� ���� ������������"
    AddNewElement "GDZSMansHave", "������������������ ������� �� �������", , , "Prop.GDZSHave", , True
    AddNewElement "GDZSMansCountWork", "������������������ ��������", , , , , False
    AddNewElement "GDZSMansCountNeed", "������������������ ���������", , , , , False
    AddNewElement "GDZSMansRezCountHave", "������������������ � ��������� ������� �������", , , , , False
    AddNewElement "GDZSMansRezCountNeed", "������������������ � ������� ���������", , , , , False
    AddNewElement "GDZSDiscr", "������������� ���������� ������������������ � ����� ������� ��������", , False
    AddNewElement "GDZSPBCount", "�� ���� �������", "{$} ({GDZSPBNeed} ���������)", , , , True
    AddNewElement "GDZSKPPCount", "��� ���� �������", "{$} ({GDZSKPPNeed} ���������)", , , , True
    AddNewElement "GDZSPBNeed", "�� ���� ���������", , , , , False
    AddNewElement "GDZSKPPNeed", "��� ���� ���������", , , , , False

    AddNewElement "Hoses38Count", "������� 38 �� ������������", , , , , False
    AddNewElement "Hoses51Count", "������� 51 �� ������������", , , , , False
    AddNewElement "Hoses66Count", "������� 66 �� ������������", , , , , False
    AddNewElement "Hoses77Count", "������� 77 �� ������������", , , , , False
    AddNewElement "Hoses89Count", "������� 89 �� ������������", , , , , False
    AddNewElement "Hoses110Count", "������� 110 �� ������������", , , , , False
    AddNewElement "Hoses150Count", "������� 150 �� ������������", , , , , False
    AddNewElement "Hoses200Count", "������� 200 �� ������������", , , , , False
    AddNewElement "Hoses250Count", "������� 250 �� ������������", , , , , False
    AddNewElement "Hoses300Count", "������� 300 �� ������������", , , , , False
    AddNewElement "HosesCount", "������� ������������ �����", "{$} (" & _
                                "{Hoses38Count} - 38 ��, {Hoses51Count} - 51 ��, {Hoses66Count} - 66 ��, {Hoses77Count} - 77 ��, {Hoses89Count} - 89 ��, " & _
                                "{Hoses110Count} - 110 ��, {Hoses150Count} - 150 ��, {Hoses200Count} - 200 ��, {Hoses250Count} - 250 ��, {Hoses300Count} - 300 ��" & _
                                ")"
    AddNewElement "HosesLength", "�������� ����� �����", "{$} �", , , , False

    AddNewElement "Hoses38Have", "������� 38 �� �������", , , "Prop.Hose38", , False
    AddNewElement "Hoses51Have", "������� 51 �� �������", , , "Prop.Hose51", , False
    AddNewElement "Hoses66Have", "������� 66 �� �������", , , "Prop.Hose66", , False
    AddNewElement "Hoses77Have", "������� 77 �� �������", , , "Prop.Hose77", , False
    AddNewElement "Hoses89Have", "������� 89 �� �������", , , "Prop.Hose89", , False
    AddNewElement "Hoses110Have", "������� 110 �� �������", , , "Prop.Hose110", , False
    AddNewElement "Hoses150Have", "������� 150 �� �������", , , "Prop.Hose150", , False
    AddNewElement "Hoses200Have", "������� 200 �� �������", , , "Prop.Hose200", , False
    AddNewElement "Hoses250Have", "������� 250 �� �������", , , "Prop.Hose250", , False
    AddNewElement "Hoses300Have", "������� 300 �� �������", , , "Prop.Hose300", , False
    AddNewElement "HosesHave", "������� ������� �����", "{$} (" & _
                                "{Hoses38Have} - 38 ��, {Hoses51Have} - 51 ��, {Hoses66Have} - 66 ��, {Hoses77Have} - 77 ��, {Hoses89Have} - 89 ��, " & _
                                "{Hoses110Have} - 110 ��, {Hoses150Have} - 150 ��, {Hoses200Have} - 200 ��, {Hoses250Have} - 250 ��, {Hoses300Have} - 300 ��" & _
                                ")", , , , True
    AddNewElement "HoseSkatka", "�������� ������� � �������", , , , , False
    
    AddNewElement "WaterValueNeed10min", "���� ��������� (�� 10���)", "{$} �"
    AddNewElement "WaterValueHave", "���� �������", "{$} �"
    AddNewElement "FoamValueeHave", "���������������� �������", "{$} �", , "Prop.Foam"

    AddNewElement "BUCount", "�� �������"
    AddNewElement "SPRCount", "��� �������"
    AddNewElement "ShtabCount", "���������� ������", , 0, , False, False
    AddNewElement "RNBDCount", "���������� �������� ����������� ������ ��������", , 0, , False, False
    AddNewElement "OchagCount", "���������� ������", , 0, , False, False
    AddNewElement "SpreadCount", "���������� ������������ ��������������� ����", , 0, , False, False
    AddNewElement "SmokeCount", "����������", , 0, , False, False
    AddNewElement "WaterSourceCount", "�������������", , 0, , False, False
    AddNewElement "DistanceCount", "����������", , 0, , False, False
    AddNewElement "DistanceMin", "���������� �� ���������� �������������", "{$} �"
    AddNewElement "LinesCount", "��������� �������� �����", , 0, , False, False
    AddNewElement "LinesPosCount", "������� ��������� �������� �����", , 0, , False, False
    AddNewElement "LinesLableCount", "������� �������� �����", , 0, , False, False
    AddNewElement "WorklinesCount", "���������� ������� �����", , 0, , False, False
    AddNewElement "MaglinesCount", "���������� ������������� �����", , 0, , False, False
    AddNewElement "BuildCount", "���������� ������", , 0, , False, False
    AddNewElement "SOCount", "���������� �������� �������������", , 0, , False, False
    AddNewElement "FireCount", "���������� �������� �������", , 0, , False, False
    AddNewElement "OrientCount", "���������", , 0, , False, False
    
    AddNewElement "InjuredCount", "������������", "{$} ���. ({DeadCount} ���. - ��������)"
    AddNewElement "DeadCount", "��������", , , , , False

    AddNewElement "AllHosesWithLabels", "��� �������� ����� � ���������", , True, , False, False
    AddNewElement "AllHosesWithPos", "��� ������� �������� ����� � ��������� ���������", , True, , False, False
    
    '��������� ����������
    AddNewElement "FireTime", "����� ������������� ������", , , "Prop.FireTime", True, False
    AddNewElement "FindTime", "����� ����������� ������", , , "Prop.FindTime", True, False
    AddNewElement "InfoTime", "����� ��������� � ������", , , "Prop.InfoTime", True, False
    AddNewElement "FirstArrivalTime", "����� �������� ������� �������������", , , "Prop.FirstArrivalTime", True, False
    AddNewElement "FirstStvolTime", "����� ������ ������� ������", , , "Prop.FirstStvolTime", True, False
    AddNewElement "LocalizationTime", "����� �����������", , , "Prop.LocalizationTime", True, False
    AddNewElement "LOGTime", "����� ���������� ��������� �������", , , "Prop.LOGTime", True, False
    AddNewElement "LPPTime", "����� ���������� ����������� ������", , , "Prop.LPPTime", True, False
    AddNewElement "FireEndTime", "����� ��������� ������", , , "Prop.FireEndTime", True, False
    AddNewElement "CurrentTime", "������� ��������������� ����� �����", , , , True, True
    AddNewElement "CurrentTimeDelta", "������� ������������� ����� �����", "�+{$}", , , True, False
    
End Sub

Private Function AddNewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_resultStrPattern As String = "", _
                        Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InTacticDataForm As Boolean = True) As Element
'������� ����� �������, ��������� ��� � ��������� ��������� � ���������� �� ���� ������
    Set AddNewElement = AddElement(NewElement(a_ID, a_CallName, a_resultStrPattern, a_startValue, a_CellID, a_InShapes, a_InTacticDataForm))
End Function

Private Function NewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_resultStrPattern As String = "", _
                        Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InTacticDataForm As Boolean = True) As Element
'������� ����� �������, ������ ��� �������� � ���������� �� ���� ������
Dim elem As Element

    Set elem = New Element
    Set NewElement = elem.SetShell(Me).Activate(a_ID, a_CallName, a_resultStrPattern, a_startValue, a_CellID, a_InShapes, a_InTacticDataForm)
End Function

Private Function AddElement(ByRef elem As Element) As Element
'��������� ������� � ��������� ��������� � ���������� �� ���� ������
Dim nextElem As Element
    
    Set nextElem = FindHigherElement(elem)
    '�������� ���������
    If nextElem Is Nothing Then
        elemCollection.Add elem, elem.ID
    Else
        elemCollection.Add elem, elem.ID, nextElem.ID
    End If
    
    '��������� ��������� � ���������� ��������
'    Debug.Assert elem.cellID = ""
'    Debug.Assert elem.cellID <> "Prop.FireTime"
    If elem.cellID <> "" Then elemWithCellsCollection.Add elem, elem.cellID
    
    Set AddElement = elem
End Function

Private Function FindHigherElement(ByRef elemIn As Element) As Element
'���������� ������� � ��������� elemCollection � ������ ������ ������ ��� � �������� (��� ������� ����� ������� ����� ����� �������� �����)
Dim elem As Element
    
'    For Each elem In elemCollection
'        If elem.callName > elemIn.callName Then
'            Set FindHigherElement = elem
'            Exit Function
'        End If
'    Next elem
End Function

'-------------------------------------��������� ��������� �������--------------------------------------------
Public Sub PrintState(Optional ByVal elemIDs As String = "")
'�������� � Debug ��� �������� � �� ��������
Dim elem As Element
Dim elemArray() As String
Dim i As Integer
    
    If elemIDs = "" Then
        For Each elem In elemCollection
            elem.PrintState
        Next elem
    Else
        elemArray = Split(elemIDs, ";")
        For i = 0 To UBound(elemArray)
            Set elem = ByID(elemArray(i))
            If Not elem Is Nothing Then elem.PrintState
        Next i
    End If
    
End Sub


'-------------------------------------��������� ������� �������----------------------------------------------
Public Function ByID(ByVal a_elemID As String) As Element
'������� ���������� ������ �� ������� � ID = a_elemID
    On Error GoTo ex
    Set ByID = elemCollection.item(a_elemID)
    
Exit Function
ex:
    Set ByID = Nothing
End Function

Public Function ByCellID(ByVal a_elemCellID As String) As Element
'������� ���������� ������ �� ������� � ������ ������ CellID = a_elemCellID
    On Error GoTo ex
    Set ByCellID = elemWithCellsCollection.item(a_elemCellID)
    
Exit Function
ex:
    Set ByCellID = Nothing
End Function

Public Function ByCallName(ByVal a_elemCallName As String) As Element
'������� ���������� ������ �� ������� � ������ ������ callName = a_elemCallName
Dim elem As Element

    For Each elem In elemCollection
        If elem.callName = a_elemCallName Then
            Set ByCallName = elem
            Exit Function
        End If
    Next elem
'���� ������ �� ������� ���������� Nothing
Set ByCallName = Nothing
End Function

Public Function GetElementsCollection(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False) As Collection
'������� ���������� ��������� ��������� ���������� �� ��������� ������
Dim coll As Collection
Dim elem As Element
Dim addFlag As Boolean
    
    Set coll = New Collection
    
    '���������� ��� �������� � ��������� ������������
    For Each elem In elemCollection
        addFlag = False
        If InStr(1, elem.ID, searchString, vbTextCompare) > 0 Then
            addFlag = True
        End If
        
        If searchByCallName Then
            If InStr(1, elem.callName, searchString) > 0 Then
                addFlag = True
            End If
        End If
        
        If addFlag Then
            coll.Add elem
        End If
    Next elem
    
Set GetElementsCollection = coll
End Function

Public Function Clear() As ElementsShell
'�������� �������� ������� � ��� ����� �������� ��������� � ��������� ���������
Dim elem As Element

    '���������� ��� �������� � �������� ������ �� ���
    For Each elem In elemCollection
        elem.Clear
    Next elem
    
    Set Clear = Me
End Function

Public Function RaiseByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String, Optional ByVal dataType As VisUnitCodes = visNumber) As ElementsShell
'����������� �������� �������� �� ����� ������
'a_CellIDs - ������������ ����� ������� � ������� ���� ����� � ���� "User.IndexPers"
Dim elem As Element
Dim cellIDArray() As String
Dim i As Integer
    
    cellIDArray = Split(a_CellIDs, ";")
    
    For i = 0 To UBound(cellIDArray)
'������� � ���������� ������ ������ ��� ������� �������� - ��� �������, �� ������ ������� ��� ������ ���������
        Set elem = ByCellID(cellIDArray(i))
        If Not elem Is Nothing Then
'            Select Case dataType
'                Case Is = visNumber
'                    elem.RaiseValue CellVal(shp, cellIDArray(i), dataType)
'                Case Is = visUnitsString
'                    elem.SetValue CellVal(shp, cellIDArray(i), dataType)
'                Case Is = visDate
'                    elem.RaiseValue CellVal(shp, cellIDArray(i), dataType)
'            End Select
            elem.RaiseValue cellVal(shp, cellIDArray(i), dataType)
        End If
    Next i
    
Set RaiseByCellID = Me
End Function

Public Function SetValByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String, Optional ByVal dataType As VisUnitCodes = visNumber) As ElementsShell
'������������� �������� �������� �� ����� ������
'a_CellIDs - ������������ ����� ������� � ������� ���� ����� � ���� "User.IndexPers"
Dim elem As Element
Dim cellIDArray() As String
Dim i As Integer
    
    cellIDArray = Split(a_CellIDs, ";")
    
    For i = 0 To UBound(cellIDArray)
'������� � ���������� ������ ������ ��� ������� �������� - ��� �������, �� ������ ������� ��� ������ ���������
        Set elem = ByCellID(cellIDArray(i))
        If Not elem Is Nothing Then
            elem.SetValue cellVal(shp, cellIDArray(i), dataType)
        End If
    Next i
    
Set SetValByCellID = Me
    
End Function

Public Function Raise(ByVal elemIDs As String, Optional ByVal val As Variant = 1) As ElementsShell
'����������� �������� ��������� �����. ������ ����������� ����� ����� � �������
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!��������! ������� � ID=" & elemArray(i) & " ����������!"
        Else
            elem.RaiseValue val
        End If
    Next i
Set Raise = Me
End Function

Public Function SetVal(ByVal elemIDs As String, Optional ByVal val As Variant = 1) As ElementsShell
'������������� �������� ��������� �����. ������ ����������� ����� ����� � �������
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!��������! ������� � ID=" & elemArray(i) & " ����������!"
        Else
            elem.SetValue val
        End If
    Next i
Set SetVal = Me
End Function

Public Function CallNames(Optional ByVal delimiter As String = ";") As String
'������� ���������� ������ ���� ��������� � ��������� ������������
Dim elem As Element
Dim tmpStr As String
    
    If elemCollection.count = 0 Then
        CallNames = ""
        Exit Function
    End If
    
    For Each elem In elemCollection
        If elem.inShapes Then
            tmpStr = tmpStr & elem.callName & delimiter
        End If
    Next elem
    
    If Len(tmpStr) = 0 Then
        CallNames = ""
        Exit Function
    End If

CallNames = Left(tmpStr, Len(tmpStr) - Len(delimiter))
End Function


'------������� ������������ �������� ���������-----------------
Public Function Result(ByVal elemID As String) As Variant
'�������� �������� �������� � ��������� elemID
Dim elem As Element
    Set elem = ByID(elemID)
    If Not elem Is Nothing Then
        Result = elem.Result
    Else
        Result = 0
    End If
End Function
Public Function ResultByCN(ByVal callName As String) As Variant
'�������� �������� �������� � ��������� CallName
Dim elem As Element
    
    Set elem = ByCallName(callName)
    
    If Not elem Is Nothing Then
        ResultByCN = elem.Result
    Else
        ResultByCN = 0
    End If
End Function

Public Function Sum(ByVal elemIDs As String) As Variant
'�������� ����� �������� ��������� � ���������� ����� ����� � ������� ������� ��������� - elemIDs
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
Dim tmp As Single
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!��������! ������� � ID=" & elemArray(i) & " ����������!"
        Else
            tmp = tmp + elem.Result
        End If
    Next i
Sum = tmp
End Function



