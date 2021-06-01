Attribute VB_Name = "m_Export"
Option Explicit

'������������ ����� � �������� � ����������
Const delimiter = " | "
'���� ������� �������
Const mChar = "$"
'���� � �������� �������� ����������
Const pathNameDonesenie = "Templates\Donesenie.dot"      '��������� � ������
Const pathNameKBD = "Templates\BD_Card.dot"      '�������� ������ ��������




'------------��������� �������� ������ �����--------------------------
Public Sub ExportToWord_Donesenie()
'������������ ������ ����� � �������� Word - � ���������
Dim wrd As Object
Dim wrdDoc As Object
Dim path As String

Dim gfsShapes As Collection
Dim gettedDate As Variant
Dim gettedCol As Collection
Dim gettedTxt As String


    '�������������� ��������� ��������� ����������� ���������� C �� ������� �:
'    fixAllGFSShapesC

    '���������� ���� � ������� ���������
    path = ThisDocument.path & pathNameDonesenie
    '������� ����� �������� Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add(path)
    wrdDoc.Activate
    
    
    
    
    
    '���� ���� ����������� �� ������� ������
    '� ������ ������ ���������� ������ ����:
'    On Error Resume Next
    
    '��������� �������� ��������� ����� ������:
    Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
    '��������� ��������:
    '---��������
'        gettedTxt = cellVal(gfsShapes, "Prop.NP_Name", visUnitsString, "")
        SetData wrd, "��", cellVal(gfsShapes, "Prop.NP_Name", visUnitsString, "")                       '���������� �����
        SetData wrd, "�����", cellVal(gfsShapes, "Prop.PersonCreate", visUnitsString, "")                      '���������, ������, �������, ���, �������� (��� �������)
        SetData wrd, "������������", cellVal(gfsShapes, "Prop.ObjectName", visUnitsString, "")          '������������
        SetData wrd, "��������������", cellVal(gfsShapes, "Prop.Affiliation", visUnitsString, "")       '�������������� �������
        SetData wrd, "�����", cellVal(gfsShapes, "Prop.Address", visUnitsString, "")                    '����� �����������
        SetData wrd, "�����������", cellVal(gfsShapes, "Prop.FireStartPlace", visUnitsString, "")              '����� ������������� ������
        SetData wrd, "���������", cellVal(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString, "")               '�������, ���, �������� (��� �������) ����, ������������� ����� � ������ ��������� � ��� � �������� ������
        SetData wrd, "���������_�", cellVal(gfsShapes, "Prop.CallPhone", visUnitsString, "")            '����� �������� ���������
        
        
        
        
'������� ���:
    '---���� � �����
        '���� ������:
        gettedDate = CDate(A.Result("FireTime"))
        If gettedDate > 0 Then
            SetData wrd, "�_����", Format(gettedDate, "DD")                             '���� ������������� ������
            SetData wrd, "�_�����", Split(Format(gettedDate, "DD MMMM"), " ")(1)        '����� ������������� ������
            SetData wrd, "�_���", Format(gettedDate, "YY")                              '����� ������������� ������
        End If
        '���� �����������:
        gettedDate = CDate(A.Result("FindTime"))
        If gettedDate > 0 Then
            SetData wrd, "���_���", Format(gettedDate, "HH")                           '��� ����������� ������
            SetData wrd, "���_���", Format(gettedDate, "NN")                           '������ ����������� ������
        End If
        '���� ���������:
        gettedDate = CDate(A.Result("InfoTime"))
        If gettedDate > 0 Then
            SetData wrd, "����_����", Format(gettedDate, "DD.MM.YYYY")                '���� ��������� � ������
            SetData wrd, "����_���", Format(gettedDate, "HH")                           '��� ��������� � ������
            SetData wrd, "����_���", Format(gettedDate, "NN")                           '������ ��������� � ������
        End If
        '����� �������� ������� �������������:
        gettedDate = CDate(A.Result("FirstArrivalTime"))
        If gettedDate > 0 Then
            SetData wrd, "1���_����", Format(gettedDate, "DD.MM.YYYY")
            SetData wrd, "1����_���", Format(gettedDate, "HH")                           '��� ��������� � ������
            SetData wrd, "1����_���", Format(gettedDate, "NN")                           '������ ��������� � ������
        End If
        
        '���� � ����� ������ ������� ������
        gettedDate = CDate(A.Result("FirstStvolTime"))
        If gettedDate > 0 Then
        SetData wrd, "1���_���", Format(gettedDate, "HH")
        SetData wrd, "1���_���", Format(gettedDate, "NN")
        End If
        
        '���� � ����� ����������� ������
        gettedDate = CDate(A.Result("LocalizationTime"))
        If gettedDate > 0 Then
        SetData wrd, "���_����", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "���_���", Format(gettedDate, "HH")
        SetData wrd, "���_���", Format(gettedDate, "NN")
        End If
        
        '���� � ����� ���������� ��������� �������
        gettedDate = CDate(A.Result("LOGTime"))
        If gettedDate > 0 Then
        SetData wrd, "���_����", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "���_���", Format(gettedDate, "HH")
        SetData wrd, "���_���", Format(gettedDate, "NN")
        End If
        
        '���� � ����� ���������� ����������� ������
        gettedDate = CDate(A.Result("LPPTime"))
        If gettedDate > 0 Then
        SetData wrd, "���_����", Format(gettedDate, "DD.MM.YYYY")
        SetData wrd, "���_���", Format(gettedDate, "HH")
        SetData wrd, "���_���", Format(gettedDate, "NN")
        End If
        
        '���������� �� ������ ��������
        Set gettedCol = A.GetGFSShapesAnd("User.IndexPers:604;Prop.SituationKind:�� ������ ��������")
        If gettedCol.Count > 0 Then
            SetData wrd, "����_����", cellVal(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
        End If
        
        '���������� ������� ����
        gettedDate = CDate(A.Result("GDZSChainsCountWork"))
        If gettedDate > 0 Then
        SetData wrd, "����_���", Format(A.Result("GDZSChainsCountWork"))
        End If
        
        '����� ���������� �������
        gettedDate = CDate(A.Result("PersonnelHave"))
        If gettedDate > 0 Then
        SetData wrd, "��_���", Format(A.Result("PersonnelHave"))
        End If
        
        
        '���������� �������� � ����������� ���������
        SetData wrd, "���_����", "�������� ��: " & A.Result("MainOverallHave") & ", ����������� ��:" & A.Result("SpecialPAHave")
        
        
        gettedTxt = ""
        gettedDate = A.Result("StvolWHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "����, "
        End If
        
        gettedDate = A.Result("StvolFoamHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "����, "
         End If
            
            gettedDate = A.Result("StvolGasHave")
        If gettedDate > 0 Then
            gettedTxt = gettedTxt & "���. "
        End If
        
        SetData wrd, "��?", gettedTxt
        
        
        '������� �����
        gettedTxt = cellVal(gfsShapes, "Prop.HumansDie", visUnitsString, "")
'        SetData wrd, "200", "������� �����: " & Split(gettedTxt, "/")(0) & "� ��� ����� �����: " & Split(gettedTxt, "/")(1) & "���������� ��: " & Split(gettedTxt, "/")(2)
        SetData wrd, "200", Split(gettedTxt, "/")(0)
        SetData wrd, "200�", Split(gettedTxt, "/")(1)
        SetData wrd, "200��", Split(gettedTxt, "/")(2)
        
        
        
         '������������ �����
        gettedTxt = cellVal(gfsShapes, "Prop.HumansInjured", visUnitsString, "")
        SetData wrd, "300", Split(gettedTxt, "/")(0)
        SetData wrd, "300�", Split(gettedTxt, "/")(1)
        SetData wrd, "300��", Split(gettedTxt, "/")(2)
        
        '���������� � �������� � ��������������
        Set gettedCol = GetVictims
        SetData wrd, "200��", gettedCol(1)
        SetData wrd, "300��", gettedCol(2)
        
        '����������/����������
        '---��������
        gettedTxt = cellVal(gfsShapes, "Prop.ConstructionsAffected", visUnitsString, "")
        SetData wrd, "�����_���", Split(gettedTxt, "/")(0)
        SetData wrd, "����_���", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellVal(gfsShapes, "Prop.FlatsAffected", visUnitsString, "")
        SetData wrd, "�����_��", Split(gettedTxt, "/")(0)
        SetData wrd, "����_��", Split(gettedTxt, "/")(1)
        '---������
        gettedTxt = cellVal(gfsShapes, "Prop.RoomsAffected", visUnitsString, "")
        SetData wrd, "�����_����", Split(gettedTxt, "/")(0)
        SetData wrd, "����_����", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellVal(gfsShapes, "Prop.SquareAffected", visUnitsString, "")
        SetData wrd, "�����_��", Split(gettedTxt, "/")(0)
        SetData wrd, "����_��", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellVal(gfsShapes, "Prop.TechnicsAffected", visUnitsString, "")
        SetData wrd, "�����_���", Split(gettedTxt, "/")(0)
        SetData wrd, "����_���", Split(gettedTxt, "/")(1)
        '---������� �������
        gettedTxt = cellVal(gfsShapes, "Prop.AgricultureAffected", visUnitsString, "")
        SetData wrd, "�����_��", ClearString(gettedTxt)
        '---������� ��������
        gettedTxt = cellVal(gfsShapes, "Prop.CattleAffected", visUnitsString, "")
        SetData wrd, "200��", ClearString(gettedTxt)
        '�������
        '---�����
        gettedTxt = cellVal(gfsShapes, "Prop.Saved", visUnitsString, "")
        SetData wrd, "����_�", Split(gettedTxt, "/")(0)
        '---�������
        SetData wrd, "����_�", Split(gettedTxt, "/")(1)
        '---����� �����
        SetData wrd, "����_��", Split(gettedTxt, "/")(2)
        
        
        '�������� �������������
        Set gettedCol = GetUniqueVals(gfsShapes, "Prop.Unit", , "-", "-")
        SetData wrd, "�������������", StrColToStr(gettedCol, ", ")
        '���, ���������� � �������������� �������� �������
        SetData wrd, "�������", GetTechniks(gettedCol)
        '���������� � ��� �������� �������
        SetData wrd, "������_��", GetReadyStringA("StvolWBHave", "������� � -", ", ") & _
                     GetReadyStringA("StvolWAHave", "������� � -", ", ") & _
                     GetReadyStringA("StvolWLHave", "�������� ������� -", ", ") & _
                     GetReadyStringA("StvolFoamHave", "������ ������� -", ", ")
                     
        
        
        
        SetData wrd, "������", cellVal(gfsShapes, "Prop.FireAutomatics", visUnitsString, "")
        SetData wrd, "���_����", "�������� ��: " & A.Result("MainOverallHave") & ", ����������� ��:" & A.Result("SpecialPAHave")
        
        
        '�������������� ����������� ����������
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesRize", visUnitsString, "")
        SetData wrd, "����", ClearString(gettedTxt)
        
        '���� � �������� ������������� ��� �������
        gettedTxt = cellVal(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "���", ClearString(gettedTxt)
        
        '�������������
        gettedTxt = cellVal(gfsShapes, "Prop.WaterSources", visUnitsString, "")
        SetData wrd, "����������", ClearString(gettedTxt)
        
        
        
'������� ���:
    '---������� ��� ������������� �������
'        clearLostMarkers wrd

        
        
End Sub


Public Sub ExportToWord_KBD()
'������������ ������ ����� � �������� Word - � �������� ������ ��������
Dim wrd As Object
Dim wrdDoc As Object
Dim path As String

Dim gfsShapes As Collection
Dim gettedDate As Variant
Dim gettedCol As Collection
Dim gettedTxt As String



    '���������� ���� � ������� ���������
    path = ThisDocument.path & pathNameKBD
    '������� ����� �������� Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add(path)
    wrdDoc.Activate
    
    
    
    
    
    '���� ���� ����������� �� ������� ������
    '� ������ ������ ���������� ������ ����:
'    On Error Resume Next
    
    '��������� �������� ��������� ����� ������:
    Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
    '��������� ��������:
        '����� �
        gettedTxt = cellVal(gfsShapes, "Prop.FireRank", visUnitsString)
        SetData wrd, "�����", gettedTxt
        '�������������
        gettedTxt = cellVal(gfsShapes, "Prop.ThisDocUnit", visUnitsString)
        SetData wrd, "����", gettedTxt
        '���� ������
        gettedDate = CDate(A.Result("FireTime"))
        SetData wrd, "�_����", Format(gettedDate, "DD.MM.YYYY")
        '������������ ����������� (�������), ��� ������������� �������������� (����� �������������), �����
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectName", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.OrgPrinadl", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.OrgPropertyType", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.Address", visUnitsString)
        SetData wrd, "���_���", gettedTxt
        '������� � �����, ���������
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectWidth", visUnitsString)
        gettedTxt = gettedTxt & "�" & cellVal(gfsShapes, "Prop.ObjectLenight", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectFloorCount", visUnitsString) & " ������"
        SetData wrd, "���_���1", gettedTxt
        '�������������� �����������, ������� ������������� ��������� ������������
        gettedTxt = cellVal(gfsShapes, "Prop.ObjectConstructions", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectSO", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.ObjectCP", visUnitsString)
        SetData wrd, "���_���2", gettedTxt
        '��� ���������� ����������� (������), ��� ��������� �����
        gettedTxt = cellVal(gfsShapes, "Prop.Guard", visUnitsString)
        gettedTxt = gettedTxt & ", " & cellVal(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString)
        SetData wrd, "���", gettedTxt
        '�����/�������
        '---������������� ������
            gettedDate = CDate(A.Result("FireTime"))
            SetData wrd, "����_��", Format(gettedDate, "HH:NN")
            SetData wrd, "����_��", "0"
        '---����������� ������
            gettedDate = CDate(A.Result("FindTime"))
            SetData wrd, "���_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "���_��", gettedTxt
        '---��������� ������
            gettedDate = CDate(A.Result("InfoTime"))
            SetData wrd, "����_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "����_��", gettedTxt
        '---������ �������
            gettedDate = DateAdd("n", 1, CDate(A.Result("InfoTime")))
            SetData wrd, "�����_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "�����_��", gettedTxt
        '---�������� �� �����
            gettedDate = CDate(A.Result("FirstArrivalTime"))
            SetData wrd, "1����_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "����_��", gettedTxt
        '---������ ������� ������
            gettedDate = CDate(A.Result("FirstStvolTime"))
            SetData wrd, "1���_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "1���_��", gettedTxt
        '---������ �������������� ���
'            gettedDate = CDate(A.Result("FireTime"))
            SetData wrd, "������_��", "---"
            SetData wrd, "������_��", "---"
        '---�����������
            gettedDate = CDate(A.Result("LocalizationTime"))
            SetData wrd, "���_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "���_��", gettedTxt
        '---���������� ��������� �������
            gettedDate = CDate(A.Result("LOGTime"))
            SetData wrd, "���_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "���_��", gettedTxt
        '---����������
            gettedDate = CDate(A.Result("LPPTime"))
            SetData wrd, "����_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "����_��", gettedTxt
        '---����������� � �����
            gettedDate = CDate(A.Result("FireEndTime"))
            SetData wrd, "����_��", Format(gettedDate, "HH:NN")
            A.Refresh Application.ActivePage.Index, gettedDate
            gettedTxt = A.Result("FireSquare")
            SetData wrd, "����_��", gettedTxt
        '���������� ������ ������� ������ ������
        Set gfsShapes = A.Refresh(Application.ActivePage.Index).gfsShapes
        '�������������
        gettedTxt = cellVal(gfsShapes, "Prop.WaterSources", visUnitsString, " ")
        SetData wrd, "�������������", ClearString(gettedTxt)
        '������� ������ ����:
        '---
            
            
            
            
        
        '���������� �� ������ ��������
        Set gettedCol = A.GetGFSShapesAnd("User.IndexPers:604;Prop.SituationKind:�� ������ ��������")
        If gettedCol.Count > 0 Then
            SetData wrd, "����", cellVal(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
        End If
        '������ ��������
        gettedTxt = GetMarkStr(gfsShapes)
        SetData wrd, "��������������", gettedTxt
        
        '����
        gettedTxt = Format(cellVal(gfsShapes, "Prop.StabCreationTime", visDate), "HH:NN") & ". " & GetStabMembers
        SetData wrd, "����", gettedTxt
        
        '��/��� - �����, ������ �������� (��������) ������� ������
        gettedTxt = GetBUSTPString
        SetData wrd, "��/���", gettedTxt
        
        '��������������, �������������� �������� ������
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesRize", visUnitsString, " ")
        SetData wrd, "����", gettedTxt
        '��������������, ����������� ����������
        gettedTxt = cellVal(gfsShapes, "Prop.CircumstancesComplicate", visUnitsString, " ")
        SetData wrd, "����", gettedTxt
        
        '���� � �������� ������������� ��� �������
        gettedTxt = cellVal(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "���", ClearString(gettedTxt)
        '� �������������� ������� ����������� (��������)
        '---�� �����������
        SetData wrd, "������", "---"
        '---� �������������� ��� � ������� ������� ������� ������� ������� �������
        SetData wrd, "�����", "---"
        
        '���� 46,90 (����, ����)
        Set gettedCol = A.GetGFSShapes("User.IndexPers:46;User.IndexPers:90")
        If gettedCol.Count > 0 Then
            SetData wrd, "����", A.Result("GDZSChainsCountWork") & " �������, " & A.Result("GDZSMansCountWork") & " ������������������"
        End If
        '1 �����
        '---�� �����������
        SetData wrd, "1��", "---"
        '2 ����� � �����
        '---�� �����������
        SetData wrd, "2��", "---"
        
        '� ������ �������� ���� ������������ ��������������
        gettedTxt = GetServicesCommunications
        SetData wrd, "�����_��", gettedTxt

        
        
        
        
        
End Sub
'Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String, _
'                    Optional ByVal ignore As Variant = 0, Optional ByVal ifIgnore As Variant = " ")
Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String)
'������ ������� � ������ ������� �� ���������� �������
    
    On Error GoTo ex
    
    '���� ��������� ������ ���������� ��������������� (��������, � ������, ���� ���� ����� 0), ������������ txt = ifIgnore
'    If txt = ignore Then txt = ifIgnore
    
    '��������� � ����� ������� ��������� �����: "markerName"=>"$markerName$"
    markerName = mChar & markerName & mChar
    
    '���������� �������� ��� ����������� ������� ������� txt
    With wrd
        .Selection.Find.ClearFormatting
        .Selection.Find.Replacement.ClearFormatting
        With .Selection.Find
            .Text = markerName
            .Replacement.Text = txt
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        .Selection.Find.Execute Replace:=2
    End With
    
ex:
    
End Sub

Private Sub clearLostMarkers(ByRef wrd As Object, Optional ByVal defaultVal As String = "     ")
Dim markers() As String
Dim marker As String
Dim i As Integer

    markers = Split("����_����;���;����������;200��;300�;300��;��;������;����_��;�������;��_�����;200��;����;����_�;����_�;����_��;�����������;�������", ";")
    
    For i = 0 To UBound(markers)
        marker = markers(i)
        SetData wrd, marker, defaultVal
    Next i
End Sub





'-------------------------������� ������������ ������� �����------------------------
'Public Function GetTechniks(ByRef gfsShapes As Collection, ByRef units As Collection) As String        '�� �������
Public Function GetTechniks(ByRef units As Collection) As String
'���������� �������������� ������ "���, ���������� � �������������� �������� �������"
Dim unit As Variant
Dim shpColl As Collection
'Dim shp As Visio.Shape
Dim tmpStr As String
Dim mainStr As String

    mainStr = ""
    For Each unit In units
        tmpStr = ""
        '��
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipAC & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "��:", ", ")
        '��
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipAL & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "��:", ", ")
        '���
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipANR & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "���:", ", ")
        '��
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipASH & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "��:", ", ")
        '���
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipASO & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "���:", ", ")
        '��
        Set shpColl = A.GetGFSShapesAnd("User.IndexPers:" & indexPers.ipKS & ";Prop.Unit:" & unit)
        tmpStr = tmpStr & GetReadyString(shpColl.Count, "��:", ", ")
        '������ �������� ��������� ���� ����������� �������� �� IndexPers
        
        '� ����� ��������� �������� ������������� (���� ���� ���� ������� ������� ���� �������)
        If tmpStr <> "" Then mainStr = mainStr & Left(tmpStr, Len(tmpStr) - 2) & "(" & unit & "); "
    Next unit
    
GetTechniks = mainStr
End Function

Private Function GetVictims() As Collection
Dim shp As Visio.Shape
Dim col As Collection
Dim deadCount As Integer
Dim casCount As Integer
Dim deads As String
Dim cased As String
Dim i As Integer
    
    Set col = A.GetGFSShapes("User.IndexPers:" & indexPers.ipPostradavshie)
    deads = " "
    cased = " "
    
    For Each shp In col
        deadCount = cellVal(shp, "Prop.CasCount")
        casCount = cellVal(shp, "Prop.iedCount")
        
        For i = 1 To 5
            If deadCount + casCount > 5 Then Exit For
            
            If i <= deadCount Then
                deads = deads & cellVal(shp, "Prop.Cas" & i, visUnitsString) & ", "
            Else
                cased = cased & cellVal(shp, "Prop.Cas" & i, visUnitsString) & ", "
            End If
        Next i
    Next shp
    
    Set GetVictims = New Collection
    GetVictims.Add deads
    GetVictims.Add cased
End Function




Private Function GetMarkStr(ByRef col As Collection) As String
'��������� ��������� ������
Dim i As Integer
Dim rowName As String
Dim marks As String
Dim shp As Visio.Shape
    
    On Error GoTo ex
    
    For Each shp In col
        For i = 0 To shp.RowCount(visSectionUser) - 1
            rowName = shp.CellsSRC(visSectionUser, i, 0).rowName
            If Len(rowName) > 9 Then
                If Left(rowName, 9) = "GFS_Info_" Then
                    '�������� � ���� ������
                    If IsGFSShapeWithIP(shp, ipDutyFace, True) Then
                        marks = marks & cellVal(shp, "Prop.Duty", visUnitsString) & " " & GetInfo(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString)) & Chr(13)
                    Else
                        marks = marks & GetInfo(shp.CellsSRC(visSectionUser, i, 0).ResultStr(visUnitsString)) & ", "
                    End If
                    
'                    cmndID = cmndID + 1
                End If
            End If
        Next i
    Next shp
    
    GetMarkStr = marks

Exit Function
ex:
    GetMarkStr = " "
End Function

Private Function GetInfo(ByVal str As String) As String
Dim strArr() As String
    
    strArr = Split(str, delimiter)
    If UBound(strArr) > 1 Then
        GetInfo = strArr(UBound(strArr))
    ElseIf UBound(strArr) = 1 Then
        GetInfo = strArr(0)
    Else
        GetInfo = " "
    End If
End Function

Private Function GetStabMembers() As String
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim gettedCol As Collection
Dim stabMember As String
    
    On Error GoTo ex
    
    Set gettedCol = A.GetGFSShapes("User.IndexPers:66")
    
    If gettedCol.Count > 0 Then
        Set shp1 = gettedCol(1)
        
        Set gettedCol = A.GetGFSShapes("User.IndexPers:65")
        For Each shp2 In gettedCol
            If shp2.SpatialRelation(shp1, 0, 0) = 4 Then
                stabMember = stabMember & cellVal(shp2, "Prop.Duty", visUnitsString) & ": " & _
                    cellVal(shp2, "Prop.FIO", visUnitsString) & ", "
            End If
        Next shp2
    End If
    
GetStabMembers = stabMember
Exit Function
ex:
    GetStabMembers = " "
End Function

Private Function GetBUSTPString() As String
Dim gettedCol As Collection
Dim shp As Visio.Shape
Dim tmpStr As String

    On Error GoTo ex

    Set gettedCol = A.GetGFSShapes("Prop.UTP_STP_Reserv:�������;Prop.UTP_STP_Reserv:������")
    
    For Each shp In gettedCol
        tmpStr = tmpStr & cellVal(shp, "User.IndexPers.Prompt", visUnitsString) & ": " & _
            "��������� - " & cellVal(shp, "Prop.NachUTP", visUnitsString) & ", " & _
            "������ - " & cellVal(shp, "Prop.UTPMission", visUnitsString) & ", " & _
            "�������� ��� - " & cellVal(shp, "Prop.UTPUnits", visUnitsString) & Chr(13)
    Next shp
    
GetBUSTPString = tmpStr
Exit Function
ex:
    GetBUSTPString = " "
End Function

Private Function GetServicesCommunications() As String
Dim gettedCol As Collection
Dim shp As Visio.Shape
Dim tmpStr As String
    
    On Error GoTo ex
    
    Set gettedCol = A.GetGFSShapes("Prop.ServiceMembership:������")
    
    For Each shp In gettedCol
        tmpStr = tmpStr & cellVal(shp, "Prop.ServiceDescription", visUnitsString) & ", "
    Next shp
    
    GetServicesCommunications = tmpStr
    
Exit Function
ex:
    GetServicesCommunications = " "
End Function




'Public Function TTT()
'Dim c As Collection
'
'    Set c = New Collection
'    c.Add "���-20"
'    c.Add "���-10"
'    c.Add "��-13"
'    c.Add "�-1"
'    Debug.Print GetTechniks(c)
''    Debug.Print A.Refresh(1).GetGFSShapesAnd("User.IndexPers:" & indexpers.ipAC & ";Prop.Unit:���-20")
'End Function



