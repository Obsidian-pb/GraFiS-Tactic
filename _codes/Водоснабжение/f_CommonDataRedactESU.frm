VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_CommonDataRedactESU 
   Caption         =   "UserForm1"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "f_CommonDataRedactESU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_CommonDataRedactESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������ ������������� ������ ��� ���
'Implements IViewTemplate

Private targetShape As Visio.Shape
Private viewTemplate As String
Private dataTemplate As DataTemplateESU





'Private Property Set IViewTemplate_dataTemplate(ByVal dt As IDataTemplate)
'    Set dataTemplate = dt
'End Property
'Private Property Get IViewTemplate_dataTemplate() As IDataTemplate
'    Set IViewTemplate_dataTemplate = dataTemplate
'End Property



'Private Function GetViewString() As String
'
'End Function


'f_CommonDataESU.FormShow "ESU:1:2:3:4:5:6:7:8:9:10:11:12:13"
'Public Sub FormShow(ByVal data As String)
''��������� ����������� �����
'Dim dt As IDataTemplate
'
'    '������� ��������� ������� ������� ������ ��� �������� ������
'    Set dt = New DataTemplateESU
'    dt.LoadData data
'
'    '��������� ���������� ������ ������ ��� ������� ������ ������ ����� � �������� DataTemplateESU
'    Set dataTemplate = dt
'
'    Me.Show
'End Sub

Public Sub FormShow(ByRef shp As Visio.Shape)
'��������� ����������� �����
'Dim dt As IDataTemplate
Dim str As String

    '��������� ������ �� ������� ������
    Set targetShape = shp
    
    '������� ������ ������
    Set dataTemplate = New DataTemplateESU
    
    '�������� �������� ������
    If shp.CellExists("Prop.Common", 0) Then
        If shp.CellExists("User.INPPWData", 0) Then
            '��������� ������� �� � ������ �����������
            If shp.Cells("User.INPPWData").ResultStr(visUnitsString) > "" Then
                '��������� ������ �� ��������� ������ ������
                dataTemplate.LoadData shp.Cells("User.INPPWData").ResultStr(visUnitsString)
            Else
                '��������� ������ �� html ��������
                dataTemplate.LoadFromHTML shp.Cells("Prop.Common").ResultStr(visUnitsString)
            End If
        End If
    End If
    
    '�������� ������ �� ������� � �������� ���������� �����
    
    
    '���������� �����
    Me.Show
End Sub





