VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WarningsForm 
   Caption         =   "��������������"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7605
   OleObjectBlob   =   "WarningsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WarningsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Win64 Then
    #If VBA7 Then
        Public FormHandle As LongPtr
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As LongPtr, _
                        ByVal nIndex As LongPtr, _
                        ByVal dwNewLong As Long) As LongPtr
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As LongPtr, _
                        ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetParent Lib "user32" ( _
                        ByVal hWndChild As LongPtr, _
                        ByVal hWndNewParent As LongPtr) As LongPtr
    #Else
        Public FormHandle As Long
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #End If
#Else
    #If VBA7 Then
        Public FormHandle As Long
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #Else
        Public FormHandle As Long
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
        Private Declare Function SetParent Lib "user32" ( _
                        ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
    #End If
#End If


Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Private Const con_BorderWidth = 6
Private Const con_BorderHeightForList = 6

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Private remarks() As String         '������ �������� ��� �����������
Const remarksItems = 28             '������� ������ �������������� (�� 0, �.�. ���������� ����� remarksItems+1)
Private remarksHided As Integer     '���������� ���������� ������� ���������

Public WithEvents menuButtonHide As CommandBarButton
Attribute menuButtonHide.VB_VarHelpID = -1
Public WithEvents menuButtonRestore As CommandBarButton
Attribute menuButtonRestore.VB_VarHelpID = -1
'Public WithEvents menuButtonOptions As CommandBarButton




'--------------------------�������� ��������� � ������� ������--------------------


Public Function Activate() As WarningsForm
    Set wAddon = ActiveWindow.Windows.Add("WarningsForm", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)

    Me.Caption = "WarningsForm"
    FormHandle = FindWindow(vbNullString, "WarningsForm")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "������ ��������"
        
    '���������� ��������� ������� ���������� ��� ������������ ��������� �����
    Set app = Visio.Application
    
    '���������� �����
    Me.Show
    
Set Activate = Me
End Function

Private Sub Stretch()
'������������� ������ ����������� ����
    Me.lstWarnings.Width = Me.Width - con_BorderWidth
    Me.lstWarnings.Height = Me.Height - con_BorderHeightForList
End Sub

Private Sub UserForm_Initialize()
    ReDim remarks(remarksItems, 1)
End Sub

'Private Sub UserForm_Terminate()
'    Set hidedRemarks = Nothing
'End Sub

Private Sub UserForm_Resize()
    Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    Set app = Nothing
    wAddon.Close
End Sub

Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
    Refresh
End Sub

'------------������ ��������������------------
Private Sub lstWarnings_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        If Y > 0 And X > 0 And Y < lstWarnings.Height And X < lstWarnings.Width Then
            DoEvents
            CreateNewMenu
        End If
    End If
End Sub




'------------��������� ���������� �����--------------------------
Public Sub Refresh()
'��������� ���������� ������ ��������������
Dim i As Integer

    On Error GoTo EX
'---�������� ������ ���������
    A.Refresh Application.ActivePage.Index
    
'---������� ����� � ������ ������� ��-���������
    Me.lstWarnings.Clear
    
'---��������� ������� ���������
    i = 0
    With A
        '����
        If remarks(i, 1) = "" Then
            If .Result("OchagCount") = 0 And (.Result("SmokeCount") > 0 Or .Result("SpreadCount") > 0 Or .Result("FireCount") > 0) Then
                remarks(i, 0) = "�� ������ ���� ������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SmokeCount") = 0 Then
                remarks(i, 0) = "�� ������� ���� ����������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SpreadCount") = 0 Then
                remarks(i, 0) = "�� ������� ���� ��������������� ������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        '����������
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BUCount") >= 3 And .Result("ShtabCount") = 0 Then
                remarks(i, 0) = "�� ������ ����������� ����"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("RNBDCount") = 0 And .Sum("OchagCount;FireCount") > 0 Then
                remarks(i, 0) = "�� ������� �������� �����������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("RNBDCount") > 1 Then
                remarks(i, 0) = "�������� ���������� ������ ���� �����"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BUCount") >= 5 And .Result("SPRCount") <= 1 Then
                remarks(i, 0) = "�� ������������ ������� ���������� �����"
            Else
                remarks(i, 0) = ""
            End If
        End If

        '����
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSPBCount") < .Result("GDZSChainsCountWork") Then
                remarks(i, 0) = "�� ���������� ����� ������������ ��� ������� ����� ���� (" & .Result("GDZSPBCount") & "/" & .Result("GDZSChainsCountWork") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsCountWork") >= 3 And .Result("GDZSKPPCount") Then
                remarks(i, 0) = "�� ������ ����������-���������� ����� ����"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSDiscr") = True Then
                remarks(i, 0) = "� ������� �������� ������ ���� ������ �������� �� ����� ��� �� ���� ������������������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And Not .options.GDZSRezRoundUp Then
                remarks(i, 0) = "������������ ��������� ������� ���� � ����������� � ������� ������� (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And .options.GDZSRezRoundUp Then
                remarks(i, 0) = "������������ ��������� ������� ���� � ����������� � ������� ������� (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        '���
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WaterSourceCount") > .Result("DistanceCount") Then
                remarks(i, 0) = "�� ������� ���������� �� ������� ������������� �� ����� ������ (" & .Result("DistanceCount") & "/" & .Result("WaterSourceCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        '������
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WorklinesCount") > .Result("LinesPosCount") Then
                remarks(i, 0) = "�� ������� ��������� (����) ��� ������ ������� ����� (" & .Result("LinesPosCount") & "/" & .Result("WorklinesCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("LinesCount") > .Result("LinesLableCount") Then
                remarks(i, 0) = "�� ������� ������� ��� ������ �������� ����� (" & .Result("LinesLableCount") & "/" & .Result("LinesCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        '���� �� ���������
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("BuildCount") > .Result("SOCount") Then
                remarks(i, 0) = "�� ������� ������� ������� ������������� ��� ������� �� ������ (" & .Result("SOCount") & "/" & .Result("BuildCount") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("OrientCount") = 0 And .Result("BuildCount") > 0 Then
                remarks(i, 0) = "�� ������� ��������� �� ���������, ����� ��� ���� ������ ��� ������� �����"
            Else
                remarks(i, 0) = ""
            End If
        End If

        '����� ��������� ������
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("FactStreamW") <> 0 And .Result("FactStreamW") < .Result("NeedStreamW") Then
                remarks(i, 0) = "������������� ����������� ������ ���� (" & .Result("FactStreamW") & " �/c < " & .Result("NeedStreamW") & " �/�)"
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("WaterValueNeed10min") > .Result("WaterValueHave") And PF_RoundUp(.Result("FactStreamW") / 32) > .Result("GetingWaterCount") Then
                remarks(i, 0) = "������������� ����� ���� ��� ������������� ������������� ������ �������"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("PersonnelHave") < .Result("PersonnelNeed") Then
                remarks(i, 0) = "������������ ������� �������, � ������ ��������� ������� (" & .Result("PersonnelHave") & "/" & .Result("PersonnelNeed") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses51Have") < .Result("Hoses51Count") Then
                remarks(i, 0) = "������������ �������� ������� 51 ��, � ������ ��������� ������� (" & .Result("Hoses51Have") & "/" & .Result("Hoses51Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses66Have") < .Result("Hoses66Count") Then
                remarks(i, 0) = "������������ �������� ������� 66 ��, � ������ ��������� ������� (" & .Result("Hoses66Have") & "/" & .Result("Hoses66Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses77Have") < .Result("Hoses77Count") Then
                remarks(i, 0) = "������������ �������� ������� 77 ��, � ������ ��������� ������� (" & .Result("Hoses77Have") & "/" & .Result("Hoses77Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses89Have") < .Result("Hoses89Count") Then
                remarks(i, 0) = "������������ �������� ������� 89 ��, � ������ ��������� ������� (" & .Result("Hoses89Have") & "/" & .Result("Hoses89Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses110Have") < .Result("Hoses110Count") Then
                remarks(i, 0) = "������������ �������� ������� 110 ��, � ������ ��������� ������� (" & .Result("Hoses110Have") & "/" & .Result("Hoses110Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses150Have") < .Result("Hoses150Count") Then
                remarks(i, 0) = "������������ �������� ������� 150 ��, � ������ ��������� ������� (" & .Result("Hoses150Have") & "/" & .Result("Hoses150Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If

        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses200Have") < .Result("Hoses200Count") Then
                remarks(i, 0) = "������������ �������� ������� 200 ��, � ������ ��������� ������� (" & .Result("Hoses200Have") & "/" & .Result("Hoses200Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses250Have") < .Result("Hoses250Count") Then
                remarks(i, 0) = "������������ �������� ������� 250 ��, � ������ ��������� ������� (" & .Result("Hoses250Have") & "/" & .Result("Hoses250Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        i = i + 1
        If remarks(i, 1) = "" Then
            If .Result("Hoses300Have") < .Result("Hoses300Count") Then
                remarks(i, 0) = "������������ �������� ������� 300 ��, � ������ ��������� ������� (" & .Result("Hoses300Have") & "/" & .Result("Hoses300Count") & ")"
            Else
                remarks(i, 0) = ""
            End If
        End If
        
        '!!!��� ���������� ����� ������� ��� ������� �� ������ ��������� ������ ������� - remarksItems!!!
    End With
    
    
    '��������� ������ ��������������
    On Error Resume Next
        lstWarnings.List = GetWarningsListArray
    On Error GoTo EX
    
    '���� �������������� �� ����������, �������� �� ����
    If lstWarnings.ListCount = 0 Then lstWarnings.AddItem "��������� �� ����������"
    
    '��������� � ����� ������ ������, ��� ����������� ����������� ������� �������
    lstWarnings.AddItem " "

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "WarningForm.Refresh"
End Sub

Private Function GetWarningsListArray() As String()
'���������� ������ ��������� ��� ���������� lstWarnings
Dim i As Integer
Dim tmpArr() As String
Dim j As Integer
Dim size As Integer

    For i = 0 To UBound(remarks, 1)
        ' ��������� ��������� ���, �� � VBA, ������ ��������� ��� ��������� ������ ���������� �������� � ����������� ������, ������� ���������� ������� �������� ������ �������� �������, � ����� ����� ���� ��� ���������((( �����
        If Not remarks(i, 0) = "" And remarks(i, 1) = "" Then
            size = size + 1
        End If
    Next i
    
    If size > 0 Then
        ReDim tmpArr(size - 1, 1)
        
        j = 0
        For i = 0 To UBound(remarks, 1)
            '���� ������� �������� � ��� ���� ���� ������� �� ���������
            If Not remarks(i, 0) = "" And remarks(i, 1) = "" Then
                tmpArr(j, 0) = remarks(i, 0)
                tmpArr(j, 1) = i
                j = j + 1
            End If
        Next i
    End If
    
GetWarningsListArray = tmpArr
End Function

'---------������� ������ � ����������� ����������� ������������
Private Sub RestoreComment()
'�������� �������� ���������� �� ����������� ���������
Dim i As Integer
    
    For i = 0 To UBound(remarks, 1)
        remarks(i, 1) = ""
    Next
    
    remarksHided = 0
End Sub

Private Sub HideComment()
'�������� ��������� �� ������� ������������
    If lstWarnings.Column(0, 0) = "��������� �� ����������" Then Exit Sub
    
    If lstWarnings.ListIndex > -1 Then
        remarks(lstWarnings.Column(1, lstWarnings.ListIndex), 1) = "h"
        remarksHided = remarksHided + 1
    End If
End Sub


'------------------������ � ����������� ����------------------
Private Sub CreateNewMenu()
'������ ����������� ���� ������� ��������
Dim popupMenuBar As CommandBar
Dim Ctrl As CommandBarControl
    
    '�������� ������ �� ����������� ����
    GetToolBar popupMenuBar, "ContextMenuListBox", msoBarPopup
    
    '������� ��������� ������ ����
    For Each Ctrl In popupMenuBar.Controls
        Ctrl.Delete
    Next
    
    '��������� ����� ������
    Set menuButtonHide = NewPopupItem(popupMenuBar, 1, 214, "�� ��������� ���������� ���������")
    Set menuButtonRestore = NewPopupItem(popupMenuBar, 1, 213, "�������� ��� ������� ���������" & " (" & remarksHided & ")", , remarksHided <> 0)
'    Set menuButtonOptions = NewPopupItem(popupMenuBar, 1, 212, "����� ���������")
    
    '���������� ����
    popupMenuBar.ShowPopup
End Sub

Private Function NewPopupItem(ByRef commBar As CommandBar, ByVal itemType As Integer, ByVal itemFace As Integer, _
ByVal itemCaption As String, Optional ByVal beginGroup As Boolean = False, Optional ByVal enableTab As Boolean = True, _
Optional itemTag As String = "") As CommandBarControl
'������� ������� ������� ������������ ���� � ���������� �� ���� ������
Dim newControl As CommandBarControl

'    On Error Resume Next
    '������� ����� �������
    Set newControl = commBar.Controls.Add(itemType)
    
    '��������� �������� ������ ��������
    With newControl
        If itemFace > 0 Then .FaceID = itemFace
        .Tag = itemTag
        .Caption = itemCaption
        .beginGroup = beginGroup
        .Enabled = enableTab
    End With
    
Set NewPopupItem = newControl
End Function

Private Sub GetToolBar(ByRef toolBar As CommandBar, ByVal toolBarName As String, ByVal barPosition As MsoBarPosition)
    On Error Resume Next
    '�������� �������� ������ �� ����������� ����
    Set toolBar = Application.CommandBars(toolBarName)

    '���� ������ ���� ���, ������� ���
    If toolBar Is Nothing Then
        Set toolBar = Application.CommandBars.Add(toolBarName, barPosition)
    End If
    
End Sub

Private Sub menuButtonHide_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    HideComment
    Refresh
End Sub

Private Sub menuButtonRestore_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    RestoreComment
    Refresh
End Sub




