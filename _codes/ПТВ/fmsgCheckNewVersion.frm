VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmsgCheckNewVersion 
   Caption         =   "�������� �� ����������"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "fmsgCheckNewVersion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmsgCheckNewVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim version As Integer
Dim description As String




Private Sub btnClose_Click()
    
    '��������� ���������
'    SaveSetting "GraFiS", "GFS_Version", "NotRemind", Me.chkNotRemind.Value
    
    '���������� �����
    Me.Hide
End Sub

Private Sub lblGoHyperlink_Click()
    Shell "cmd /cstart " & lblGoHyperlink.Tag
End Sub

Public Sub CheckUpdates()
    
Dim lastCheckDate As Date
Dim currentVersion As Integer
    

    
    '��������� ���� ��������� �������� - ���� ������ ������ ����� �������
    lastCheckDate = CDate(GetSetting("GraFiS", "GFS_Version", "LastCheckDate", Now()))
    If DateDiff("d", lastCheckDate, Now()) >= 1 Then
        '�������� ������� ������ �������. ���� �������� � ������ - �������
        If Not GetVersion(currentVersion) Then
             SaveSetting "GraFiS", "GFS_Version", "LastCheckDate", CStr(Now())
             Exit Sub
        End If
        
        GetData currentVersion      '��.�.��  - 1 ������� ������, 2 ������������� ������, 3 - ������� ������
        SaveSetting "GraFiS", "GFS_Version", "LastCheckDate", CStr(Now())
    End If
    

End Sub

Private Function GetData(ByVal a_version As Integer) As Boolean
'������� ���������� ������ ���� ������� ����� �� �������, � ��� �� ����������� ��� �������
Dim xmlDoc As MSXML2.DOMDocument60
Dim xmlNode As MSXML2.IXMLDOMNode

    GetData = False
    
    On Error GoTo EX
    
    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.Load "http://graphicalfiresets.ru/Source/GFSVersion.xml"
    
    Set xmlNode = xmlDoc.SelectSingleNode("//version")
    version = CInt(xmlNode.Text)
    Set xmlNode = xmlDoc.SelectSingleNode("//description")
    description = xmlNode.Text
    
    If a_version < version Then
        Me.txtVersion = VersionToMaskString(version)
        Me.txtDescription = description
        Me.Caption = "�������� �� ���������� (������� ������ " & VersionToMaskString(a_version) & ")"
        Me.Show
        
        GetData = True
    End If
    

    Set xmlNode = Nothing
    Set xmlDoc = Nothing
Exit Function
EX:
    GetData = False
End Function



Private Function VersionToMaskString(ByVal a_version As Integer) As String
Dim firstString As String
Dim secondString As String

    firstString = CStr(a_version)
    secondString = Left(firstString, 2) & "." & Mid(firstString, 3, 1) & "." & Right(firstString, 2)
    
    VersionToMaskString = secondString
    
End Function

Private Function GetVersion(ByRef version As Integer) As Boolean
'������� ���������� ������� ������ ������ ��������� � ����� Version.txt
Dim txt As String
Dim S_FilePath As String

On Error GoTo Tail

    S_FilePath = ThisDocument.path & "\Version.txt"
    
    Open S_FilePath For Input As #1 ' ��������� ���� Version.txt ��� ������

    Line Input #1, txt ' ������ 1 ������ ������
    version = CInt(txt)
    GetVersion = True
    
    Close #1
Exit Function

Tail:
MsgBox "���-�� ����� �� ���! ��������� ������� ����� Version.txt (���� ������ ���������� � ��� �� ��������, ��� � ������� ����.", vbCritical
GetVersion = False
End Function

