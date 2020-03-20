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

Private remarks(27) As Boolean '������ ���������� ��� ����� ����������� ���������
Private remarksHided As Integer '���������� ���������� ������� ���������

Public WithEvents menuButtonHide As CommandBarButton
Attribute menuButtonHide.VB_VarHelpID = -1
Public WithEvents menuButtonRestore As CommandBarButton
Attribute menuButtonRestore.VB_VarHelpID = -1
Public WithEvents menuButtonOptions As CommandBarButton
Attribute menuButtonOptions.VB_VarHelpID = -1




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



'------------��������� ���������� �����--------------------------
Public Sub Refresh()
'��������� ���������� ������ ��������������
Dim Comment As Boolean
Dim i As Integer

'---�������� ������ ���������
'    Set elemCollection = A.Refresh(Application.ActivePage.Index).elements.GetElementsCollection("")
    A.Refresh Application.ActivePage.Index
    
'---������� ����� � ������ ������� ��-���������
    Me.lstWarnings.Clear
    
    Comment = False
    remarksHided = 0
    
'---��������� ������� ���������
    With A
        '����
        If remarks(0) = False Then
            If .Result("OchagCount") = 0 And (.Result("SmokeCount") > 0 Or .Result("SpreadCount") > 0 Or .Result("FireCount") > 0) Then
                lstWarnings.AddItem "�� ������ ���� ������"
                Comment = True
            End If
        End If
        
        If remarks(1) = False Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SmokeCount") = 0 Then
                lstWarnings.AddItem "�� ������� ���� ����������"
                Comment = True
            End If
        End If

        If remarks(2) = False Then
            If .Sum("OchagCount;FireCount") > 0 And .Result("SpreadCount") = 0 Then
                lstWarnings.AddItem "�� ������� ���� ��������������� ������"
                Comment = True
            End If
        End If

        '����������
        If remarks(3) = False Then
            If .Result("BUCount") >= 3 And .Result("ShtabCount") = 0 Then
                lstWarnings.AddItem "�� ������ ����������� ����"
                Comment = True
            End If
        End If

        If remarks(4) = False Then
'            If vOC_InfoAnalizer.pi_RNBDCount = 0 And vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 Then
            If .Result("RNBDCount") = 0 And .Sum("OchagCount;FireCount") > 0 Then
                lstWarnings.AddItem "�� ������� �������� �����������"
                Comment = True
            End If
        End If

        If remarks(5) = False Then
'            If vOC_InfoAnalizer.pi_RNBDCount > 1 Then
            If .Result("RNBDCount") > 1 Then
                lstWarnings.AddItem "�������� ���������� ������ ���� �����"
                Comment = True
            End If
        End If

        If remarks(6) = False Then
'            If vOC_InfoAnalizer.pi_BUCount >= 5 And vOC_InfoAnalizer.pi_SPRCount <= 1 Then
            If .Result("BUCount") >= 5 And .Result("SPRCount") <= 1 Then
                lstWarnings.AddItem "�� ������������ ������� ���������� �����"
                Comment = True
            End If
        End If

        '����
        If remarks(7) = False Then
'            If vOC_InfoAnalizer.pi_GDZSpbCount < vOC_InfoAnalizer.pi_GDZSChainsCount Then
            If .Result("GDZSPBCount") < .Result("GDZSChainsCountWork") Then
                lstWarnings.AddItem "�� ���������� ����� ������������ ��� ������� ����� ���� (" & .Result("GDZSPBCount") & "/" & .Result("GDZSChainsCountWork") & ")"
                Comment = True
            End If
        End If

        If remarks(8) = False Then
'            If vOC_InfoAnalizer.pi_GDZSChainsCount >= 3 And vOC_InfoAnalizer.pi_KPPCount = 0 Then
            If .Result("GDZSChainsCountWork") >= 3 And .Result("GDZSKPPCount") Then
                lstWarnings.AddItem "�� ������ ����������-���������� ����� ����"
                Comment = True
            End If
        End If

        If remarks(9) = False Then
'            If vOC_InfoAnalizer.pb_GDZSDiscr = True Then
            If .Result("GDZSDiscr") = True Then
                lstWarnings.AddItem "� ������� �������� ������ ���� ������ �������� �� ����� ��� �� ���� ������������������"
                Comment = True
            End If
        End If

        If remarks(10) = False Then
'            If Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) > vOC_InfoAnalizer.pi_GDZSChainsRezCount And bo_GDZSRezRoundUp = False Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And Not .options.GDZSRezRoundUp Then
                    lstWarnings.AddItem "������������ ��������� ������� ���� � ����������� � ������� ������� (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
                    Comment = True
            End If
        End If

        If remarks(10) = False Then
'            If Fix((vOC_InfoAnalizer.ps_GDZSChainsRezNeed / 0.3334) * 0.3333) + 1 > vOC_InfoAnalizer.pi_GDZSChainsRezCount And bo_GDZSRezRoundUp = True And vOC_InfoAnalizer.pi_GDZSChainsCount <> 0 Then
            If .Result("GDZSChainsRezCountNeed") > .Result("GDZSChainsRezCountHave") And .options.GDZSRezRoundUp Then
                    lstWarnings.AddItem "������������ ��������� ������� ���� � ����������� � ������� ������� (" & .Result("GDZSChainsRezCountHave") & "/" & .Result("GDZSChainsRezCountNeed") & ")"
                    Comment = True
            End If
        End If

        '���
        If remarks(11) = False Then
'            If vOC_InfoAnalizer.pi_WaterSourceCount > vOC_InfoAnalizer.pi_distanceCount Then
            If .Result("WaterSourceCount") > .Result("DistanceCount") Then
                lstWarnings.AddItem "�� ������� ���������� �� ������� ������������� �� ����� ������ (" & .Result("DistanceCount") & "/" & .Result("WaterSourceCount") & ")"
                Comment = True
            End If
        End If

        '������
        If remarks(12) = False Then
'            If vOC_InfoAnalizer.pi_WorklinesCount > vOC_InfoAnalizer.pi_linesPosCount Then
            If .Result("WorklinesCount") > .Result("LinesPosCount") Then
                lstWarnings.AddItem "�� ������� ��������� (����) ��� ������ ������� ����� (" & .Result("LinesPosCount") & "/" & .Result("WorklinesCount") & ")"
                Comment = True
            End If
        End If

        If remarks(13) = False Then
'            If vOC_InfoAnalizer.pi_linesCount > vOC_InfoAnalizer.pi_linesLableCount Then
            If .Result("LinesCount") > .Result("LinesLableCount") Then
                lstWarnings.AddItem "�� ������� �������� ��� ������ �������� ����� (" & .Result("LinesLableCount") & "/" & .Result("LinesCount") & ")"
                Comment = True
            End If
        End If

        '���� �� ���������
        If remarks(14) = False Then
'            If vOC_InfoAnalizer.pi_BuildCount > vOC_InfoAnalizer.pi_SOCount Then
            If .Result("BuildCount") > .Result("SOCount") Then
                lstWarnings.AddItem "�� ������� ������� ������� ������������� ��� ������� �� ������ (" & .Result("SOCount") & "/" & .Result("BuildCount") & ")"
                Comment = True
            End If
        End If

        If remarks(15) = False Then
'            If vOC_InfoAnalizer.pi_OrientCount = 0 And vOC_InfoAnalizer.pi_BuildCount > 0 Then
            If .Result("OrientCount") = 0 And .Result("BuildCount") > 0 Then
                lstWarnings.AddItem "�� ������� ��������� �� ���������, ����� ��� ���� ������ ��� ������� �����"
                Comment = True
            End If
        End If

        '����� ��������� ������
        If remarks(16) = False Then
'            If vOC_InfoAnalizer.ps_FactStreemW <> 0 And vOC_InfoAnalizer.ps_FactStreemW < vOC_InfoAnalizer.ps_NeedStreemW Then
            If .Result("FactStreamW") <> 0 And .Result("FactStreamW") < .Result("NeedStreamW") Then
                lstWarnings.AddItem "������������� ����������� ������ ���� (" & .Result("FactStreamW") & " �/c < " & .Result("NeedStreamW") & " �/�)"
                Comment = True
            End If
        End If

        If remarks(17) = False Then
'            If (vOC_InfoAnalizer.ps_FactStreemW * 600) > vOC_InfoAnalizer.pi_WaterValueHave Then
            If .Result("WaterValueNeed10min") > .Result("WaterValueHave") And PF_RoundUp(.Result("FactStreamW") / 32) > .Result("GetingWaterCount") Then
'                If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) > vOC_InfoAnalizer.pi_GetingWaterCount Then lstWarnings.AddItem "������������� ������������� ������ �������" '& (" & vOC_InfoAnalizer.pi_GetingWaterCount & "/" & PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) & ")"
                lstWarnings.AddItem "������������� ����� ���� ��� ������������� ������������� ������ �������"
                Comment = True
            End If
        End If

        If remarks(18) = False Then
'            If vOC_InfoAnalizer.pi_PersonnelHave < vOC_InfoAnalizer.pi_PersonnelNeed Then
            If .Result("PersonnelHave") < .Result("PersonnelNeed") Then
                lstWarnings.AddItem "������������ ������� �������, � ������ ��������� ������� (" & .Result("PersonnelHave") & "/" & .Result("PersonnelNeed") & ")"
                Comment = True
            End If
        End If

        If remarks(19) = False Then
'            If vOC_InfoAnalizer.pi_Hoses51Have < vOC_InfoAnalizer.pi_Hoses51Count Then
            If .Result("Hoses51Have") < .Result("Hoses51Count") Then
                lstWarnings.AddItem "������������ �������� ������� 51 ��, � ������ ��������� ������� (" & .Result("Hoses51Have") & "/" & .Result("Hoses51Count") & ")"
                Comment = True
            End If
        End If

        If remarks(20) = False Then
'            If vOC_InfoAnalizer.pi_Hoses66Have < vOC_InfoAnalizer.pi_Hoses66Count Then
            If .Result("Hoses66Have") < .Result("Hoses66Count") Then
                lstWarnings.AddItem "������������ �������� ������� 66 ��, � ������ ��������� ������� (" & .Result("Hoses66Have") & "/" & .Result("Hoses66Count") & ")"
                Comment = True
            End If
        End If

        If remarks(21) = False Then
'            If vOC_InfoAnalizer.pi_Hoses77Have < vOC_InfoAnalizer.pi_Hoses77Count Then
            If .Result("Hoses77Have") < .Result("Hoses77Count") Then
                lstWarnings.AddItem "������������ �������� ������� 77 ��, � ������ ��������� ������� (" & .Result("Hoses77Have") & "/" & .Result("Hoses77Count") & ")"
                Comment = True
            End If
        End If

        If remarks(22) = False Then
            If .Result("Hoses89Have") < .Result("Hoses89Count") Then
                lstWarnings.AddItem "������������ �������� ������� 89 ��, � ������ ��������� ������� (" & .Result("Hoses89Have") & "/" & .Result("Hoses89Count") & ")"
                Comment = True
            End If
        End If

        If remarks(23) = False Then
            If .Result("Hoses110Have") < .Result("Hoses110Count") Then
                lstWarnings.AddItem "������������ �������� ������� 110 ��, � ������ ��������� ������� (" & .Result("Hoses110Have") & "/" & .Result("Hoses110Count") & ")"
                Comment = True
            End If
        End If

        If remarks(24) = False Then
            If .Result("Hoses150Have") < .Result("Hoses150Count") Then
                lstWarnings.AddItem "������������ �������� ������� 150 ��, � ������ ��������� ������� (" & .Result("Hoses150Have") & "/" & .Result("Hoses150Count") & ")"
                Comment = True
            End If
        End If

        If remarks(25) = False Then
            If .Result("Hoses200Have") < .Result("Hoses200Count") Then
                lstWarnings.AddItem "������������ �������� ������� 200 ��, � ������ ��������� ������� (" & .Result("Hoses200Have") & "/" & .Result("Hoses200Count") & ")"
                Comment = True
            End If
        End If

        If remarks(26) = False Then
            If .Result("Hoses250Have") < .Result("Hoses250Count") Then
                lstWarnings.AddItem "������������ �������� ������� 250 ��, � ������ ��������� ������� (" & .Result("Hoses250Have") & "/" & .Result("Hoses250Count") & ")"
                Comment = True
            End If
        End If

        If remarks(27) = False Then
            If .Result("Hoses300Have") < .Result("Hoses300Count") Then
                lstWarnings.AddItem "������������ �������� ������� 300 ��, � ������ ��������� ������� (" & .Result("Hoses300Have") & "/" & .Result("Hoses300Count") & ")"
                Comment = True
            End If
        End If
    End With
    If Comment = False Then lstWarnings.AddItem "��������� �� ����������"
    
    '��������� � ����� ������ ������, ��� ����������� ����������� ������� �������
    lstWarnings.AddItem " "

'������� ������� ���������
    For i = 0 To UBound(remarks)
        If remarks(i) = True Then remarksHided = remarksHided + 1
    Next
    
End Sub

