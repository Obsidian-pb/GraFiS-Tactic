Attribute VB_Name = "m_Export"
Option Explicit

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
        SetData wrd, "��", cellval(gfsShapes, "Prop.NP_Name", visUnitsString, "")                       '���������� �����
        SetData wrd, "�����", cellval(gfsShapes, "Prop.PersonCreate", visUnitsString, "")                      '���������, ������, �������, ���, �������� (��� �������)
        SetData wrd, "������������", cellval(gfsShapes, "Prop.ObjectName", visUnitsString, "")          '������������
        SetData wrd, "��������������", cellval(gfsShapes, "Prop.Affiliation", visUnitsString, "")       '�������������� �������
        SetData wrd, "�����", cellval(gfsShapes, "Prop.Address", visUnitsString, "")                    '����� �����������
        SetData wrd, "�����������", cellval(gfsShapes, "Prop.FireStartPlace", visUnitsString, "")              '����� ������������� ������
        SetData wrd, "���������", cellval(gfsShapes, "Prop.CallerFIOAndCase", visUnitsString, "")               '�������, ���, �������� (��� �������) ����, ������������� ����� � ������ ��������� � ��� � �������� ������
        SetData wrd, "���������_�", cellval(gfsShapes, "Prop.CallPhone", visUnitsString, "")            '����� �������� ���������
        
        
        
        
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
            SetData wrd, "����_����", cellval(gettedCol, "Prop.SituationDescription", visUnitsString, " ")
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
        gettedTxt = cellval(gfsShapes, "Prop.HumansDie", visUnitsString, "")
'        SetData wrd, "200", "������� �����: " & Split(gettedTxt, "/")(0) & "� ��� ����� �����: " & Split(gettedTxt, "/")(1) & "���������� ��: " & Split(gettedTxt, "/")(2)
        SetData wrd, "200", Split(gettedTxt, "/")(0)
        SetData wrd, "200�", Split(gettedTxt, "/")(1)
        SetData wrd, "200��", Split(gettedTxt, "/")(2)
        
        
        
         '������������ �����
        gettedTxt = cellval(gfsShapes, "Prop.HumansInjured", visUnitsString, "")
        SetData wrd, "300", Split(gettedTxt, "/")(0)
        SetData wrd, "300�", Split(gettedTxt, "/")(1)
        SetData wrd, "300��", Split(gettedTxt, "/")(2)
        
        '���������� � �������� � ��������������
        Set gettedCol = GetVictims
        SetData wrd, "200��", gettedCol(1)
        SetData wrd, "300��", gettedCol(2)
        
        '����������/����������
        '---��������
        gettedTxt = cellval(gfsShapes, "Prop.ConstructionsAffected", visUnitsString, "")
        SetData wrd, "�����_���", Split(gettedTxt, "/")(0)
        SetData wrd, "����_���", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellval(gfsShapes, "Prop.FlatsAffected", visUnitsString, "")
        SetData wrd, "�����_��", Split(gettedTxt, "/")(0)
        SetData wrd, "����_��", Split(gettedTxt, "/")(1)
        '---������
        gettedTxt = cellval(gfsShapes, "Prop.RoomsAffected", visUnitsString, "")
        SetData wrd, "�����_����", Split(gettedTxt, "/")(0)
        SetData wrd, "����_����", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellval(gfsShapes, "Prop.SquareAffected", visUnitsString, "")
        SetData wrd, "�����_��", Split(gettedTxt, "/")(0)
        SetData wrd, "����_��", Split(gettedTxt, "/")(1)
        '---�������
        gettedTxt = cellval(gfsShapes, "Prop.TechnicsAffected", visUnitsString, "")
        SetData wrd, "�����_���", Split(gettedTxt, "/")(0)
        SetData wrd, "����_���", Split(gettedTxt, "/")(1)
        '---������� �������
        gettedTxt = cellval(gfsShapes, "Prop.AgricultureAffected", visUnitsString, "")
        SetData wrd, "�����_��", ClearString(gettedTxt)
        '---������� ��������
        gettedTxt = cellval(gfsShapes, "Prop.CattleAffected", visUnitsString, "")
        SetData wrd, "200��", ClearString(gettedTxt)
        '�������
        '---�����
        gettedTxt = cellval(gfsShapes, "Prop.Saved", visUnitsString, "")
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
                     
        
        
        
        SetData wrd, "������", cellval(gfsShapes, "Prop.FireAutomatics", visUnitsString, "")
        SetData wrd, "���_����", "�������� ��: " & A.Result("MainOverallHave") & ", ����������� ��:" & A.Result("SpecialPAHave")
        
        
        '�������������� ����������� ����������
        gettedTxt = cellval(gfsShapes, "Prop.CircumstancesRize", visUnitsString, "")
        SetData wrd, "����", ClearString(gettedTxt)
        
        '���� � �������� ������������� ��� �������
        gettedTxt = cellval(gfsShapes, "Prop.SiS", visUnitsString, "")
        SetData wrd, "���", ClearString(gettedTxt)
        
        '�������������
        gettedTxt = cellval(gfsShapes, "Prop.WaterSources", visUnitsString, "")
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
    '---���� � �����
        '���� ������:
        gettedDate = CDate(A.Result("FireTime"))
        SetData wrd, "�_����", Format(gettedDate, "DD.MM.YYYY")



'        gettedDate = CDate(A.Result("FireTime"))
'        SetData wrd, "�_����", Format(gettedDate, "DD")                             '���� ������������� ������
'        SetData wrd, "�_�����", Split(Format(gettedDate, "DD MMMM"), " ")(1)        '����� ������������� ������
'        SetData wrd, "�_���", Format(gettedDate, "YY")                              '����� ������������� ������
'        '���� ���������:
'        gettedDate = CDate(A.Result("InfoTime"))
'        SetData wrd, "����_����", Format(gettedDate, "DD MMMM YYYY")                '���� ��������� � ������
'        SetData wrd, "����_���", Format(gettedDate, "HH")                           '��� ��������� � ������
'        SetData wrd, "����_���", Format(gettedDate, "NN")                           '������ ��������� � ������
'        '����� �������� ������� �������������:
'        gettedDate = CDate(A.Result("FirstArrivalTime"))
'        SetData wrd, "1����_���", Format(gettedDate, "HH")                           '��� ��������� � ������
'        SetData wrd, "1����_���", Format(gettedDate, "NN")                           '������ ��������� � ������
'
'
'        '�������� �������������
'        Set gettedCol = GetUniqueVals(gfsShapes, "Prop.Unit", , "-", "-")
'        SetData wrd, "�������������", StrColToStr(gettedCol, ", ")
'        '���, ���������� � �������������� �������� �������
'        SetData wrd, "�������", GetTechniks(gettedCol)
        
        
        
        

        
        
End Sub
'Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String, _
'                    Optional ByVal ignore As Variant = 0, Optional ByVal ifIgnore As Variant = " ")
Private Sub SetData(ByRef wrd As Object, ByVal markerName As String, ByVal txt As String)
'������ ������� � ������ ������� �� ���������� �������
    
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
        deadCount = cellval(shp, "Prop.CasCount")
        casCount = cellval(shp, "Prop.iedCount")
        
        For i = 1 To 5
            If deadCount + casCount > 5 Then Exit For
            
            If i <= deadCount Then
                deads = deads & cellval(shp, "Prop.Cas" & i, visUnitsString) & ", "
            Else
                cased = cased & cellval(shp, "Prop.Cas" & i, visUnitsString) & ", "
            End If
        Next i
    Next shp
    
    Set GetVictims = New Collection
    GetVictims.Add deads
    GetVictims.Add cased
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



