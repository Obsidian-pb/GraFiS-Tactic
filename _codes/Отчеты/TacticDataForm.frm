VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TacticDataForm 
   Caption         =   "����������� ������"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   OleObjectBlob   =   "TacticDataForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TacticDataForm"
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

Public WithEvents menuButtonHide As CommandBarButton
Attribute menuButtonHide.VB_VarHelpID = -1
Public WithEvents menuButtonRestore As CommandBarButton
Attribute menuButtonRestore.VB_VarHelpID = -1
Public WithEvents menuButtonOptions As CommandBarButton
Attribute menuButtonOptions.VB_VarHelpID = -1

Private curShapeID As Long
Private lastCalcTime As Double
Const recalcInterval = 0.0000116

'Private elemCollection As Collection

'--------------------------�������� ��������� � ������� ������--------------------
Public Function Activate() As TacticDataForm
    '����� ���������� �� �����������! ���� ������ ����� ������� �� ��������, �� ��� ����� ����������� ������, ����� - ���. ����� ���� ��������� ������������ �����
    If WarningsForm.Visible = True Then
        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom + visWSAnchorMerged, visAnchorBarAddon, , , 600, , "MC", "MC")
    Else
        Set wAddon = ActiveWindow.Windows.Add("TacticDataForm", visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 600, 210, "MC", "MC")
    End If
    
    Me.Caption = "TacticDataForm"
    FormHandle = FindWindow(vbNullString, "TacticDataForm")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "������ ����������� ������"
        
    '���������� ��������� ������� ���������� ��� ������������ ��������� �����
    Set app = Visio.Application
    
    '���������� �����
    Me.Show
    
Set Activate = Me
End Function

Private Sub Stretch()
'������������� ������ ����������� ����
    Me.lstTacticData.Width = Me.Width - con_BorderWidth
    Me.lstTacticData.Height = Me.Height - con_BorderHeightForList
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
    
    If curShapeID <> Cell.Shape.ID Then
        Refresh
        curShapeID = Cell.Shape.ID
        'Debug.Print Cell.Shape.Name
    Else
        If lastCalcTime + recalcInterval < Now() Then
            Refresh
            'Debug.Print Cell.Shape.Name
        End If
    End If
    
    lastCalcTime = Now()
End Sub



'------------��������� ���������� �����--------------------------
Public Sub Refresh()
'��������� ���������� ������ ��������������
Dim i As Integer
Dim elemCollection As Collection
Dim elem As Element

'---�������� ������ ���������
    A.Refresh Application.ActivePage.Index
    
'---������� ����� � ������ ������� ��-���������
    Me.lstTacticData.Clear

'---��������� ������� ���������
    Set elemCollection = A.elements.GetElementsCollection("")
    
    i = 0
'    With A
    For Each elem In elemCollection
        If elem.inTacticForm Then
            If elem.Result > 0 Then
                lstTacticData.AddItem elem.callName, i
                lstTacticData.List(i, 1) = elem.ResultStr
                i = i + 1
            End If
        End If
    Next elem
'    End With
    
    '��������� � ����� ������ ������, ��� ����������� ����������� ������� �������
    lstTacticData.AddItem " "

End Sub







