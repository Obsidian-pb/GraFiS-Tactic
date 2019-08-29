VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertiesForm 
   Caption         =   "������ ������"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   OleObjectBlob   =   "PropertiesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PropertiesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CBut_OK_Click()
Me.Hide
sp_DataRefresh
End Sub

Private Sub UserForm_Activate()
'�������� ����� ��� ��������� � �������������� ������� ������� ������
'---��������� ����������
Dim vpVS_DocShape As Visio.Shape

'---���������� ������ ����-����� ���������
    Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---�������� ������ �� ����-����� ���������
    Me.TB_City = vpVS_DocShape.Cells("User.City").ResultStr(Visio.visNone)
    Me.TB_Adress = vpVS_DocShape.Cells("User.Adress").ResultStr(Visio.visNone)
    Me.CB_FireRating = vpVS_DocShape.Cells("User.FireRating").ResultStr(Visio.visNone)
    Me.CB_Object = vpVS_DocShape.Cells("User.Object").ResultStr(Visio.visNone)

End Sub



Private Sub UserForm_Initialize()
'---��������� ������
'---��������� ����������
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String

    On Error GoTo EX
    '---������ �������� �������������
    For i = 1 To 5
        Me.CB_FireRating.AddItem (i)
    Next i
    
    '---������ �������� �������������
    '---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT ��������, [���������] " & _
        "FROM �_������������� " & _
        "WHERE (([���������])='������ � ����������')" & _
        "ORDER BY �_�������������.��������;"

    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
        
    '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
    With rst
        .MoveFirst
        Do Until .EOF
            Me.CB_Object.AddItem (![��������])
            .MoveNext
        Loop
    End With

Exit Sub
EX:
    SaveLog Err, "UserForm_Initialize"
End Sub

Private Sub UserForm_Terminate()
'������� ����� � ���������� ������
sp_DataRefresh
End Sub

Private Sub sp_DataRefresh()
'---��������� ����������
Dim vpVS_DocShape As Visio.Shape

'---���������� ������ ����-����� ���������
Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---�������� ������ �� ����-����� ���������
vpVS_DocShape.Cells("User.City").FormulaU = Chr(34) & Me.TB_City.Value & Chr(34)
vpVS_DocShape.Cells("User.Adress").FormulaU = Chr(34) & Me.TB_Adress.Value & Chr(34)
vpVS_DocShape.Cells("User.FireRating").FormulaU = Chr(34) & Me.CB_FireRating.Value & Chr(34)
vpVS_DocShape.Cells("User.Object").FormulaU = Chr(34) & Me.CB_Object.Value & Chr(34)
End Sub
