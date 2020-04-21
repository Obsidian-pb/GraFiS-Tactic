VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportJPG 
   Caption         =   "������� ��������"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   OleObjectBlob   =   "ExportJPG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportJPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBClose_Click()
    ExportJPG.Hide
End Sub


Private Sub UserForm_Activate()

    '������������� ����������� ������������, � ������, ���� ��  �������
    TryToHideExportBar

ExportJPG.Repaint

Dim curPath As String
Dim curPage As String
Dim i As Integer
Dim progressWidthStep As Double



    curPath = Application.ActiveDocument.path
    
    If curPath = "" Then
        MsgBox "������� ���������� ��������� ������� ��������"
        Exit Sub
    End If
    
    '������ ��������� �������� ��� ������ �������
    progressWidthStep = Me.ProgressFrame.Width / Application.ActiveDocument.Pages.Count
    ProgressFillFrame.Width = 0
    
    '���������� ��� �������� � ��������� ��
    For i = 1 To Application.ActiveDocument.Pages.Count
        Application.ActiveWindow.Page = Application.ActiveDocument.Pages(i)
        curPage = Application.ActivePage.Name
        Application.ActiveWindow.Page.Export curPath & curPage & ".jpg"
    
'        ExportProgress.Value = i
        ProgressFillFrame.Width = ProgressFillFrame.Width + progressWidthStep
    Next i


Label2.Visible = True
CBClose.Enabled = True

End Sub

Private Sub TryToHideExportBar()
    On Error Resume Next
    ExportProgress.Visible = False
End Sub
