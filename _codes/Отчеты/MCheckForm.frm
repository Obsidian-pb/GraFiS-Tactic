VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MCheckForm 
   Caption         =   " Мастер проверок схемы - Бета-версия"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8628
   OleObjectBlob   =   "MCheckForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, _
ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
 Const MOUSEEVENTF_ABSOLUTE = &H8000
 Const MOUSEEVENTF_LEFTDOWN = &H2
 Const MOUSEEVENTF_LEFTUP = &H4
 Const MOUSEEVENTF_MIDDLEDOWN = &H20
 Const MOUSEEVENTF_MIDDLEUP = &H40
 Const MOUSEEVENTF_MOVE = &H1
 Const MOUSEEVENTF_RIGHTDOWN = &H8
 Const MOUSEEVENTF_RIGHTUP = &H10

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
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private f_MCheckForm As MCheckForm
Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Private Const con_BorderWidth = 6
Private Const con_BorderHeightForList = 20



Private Sub ListBox1_Click()
    MasterCheckRefresh
End Sub


 
Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal x As Single, ByVal y As Single)
    If Button = 2 Then
        If y > 0 And x > 0 And y < ListBox1.Height And x < ListBox1.Width Then
            mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&
            mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
            DoEvents
            CreateNewMenu
        End If
    End If
 End Sub

Private Sub ListBox2_Click()
    MasterCheckRefresh
End Sub

Private Sub UserForm_Activate()
    Set wAddon = ActiveWindow.Windows.Add("MCheck", visWSVisible + visWSDockedBottom, visAnchorBarAddon, , , 300, 210)

    Me.Caption = "MCheck"
    FormHandle = FindWindow(vbNullString, "MCheck")
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = "Мастер проверок"
    

    
    Set app = Visio.Application
End Sub

Private Sub UserForm_Click()
    MasterCheckRefresh
End Sub

Private Sub pS_Stretch()
'Устанавливаем размер содержимого окна
    Me.MultiPage1.Width = Me.InsideWidth - con_BorderWidth
    Me.MultiPage1.Height = Me.InsideHeight - con_BorderWidth
    Me.ListBox1.Width = Me.MultiPage1.Width - con_BorderWidth
    Me.ListBox1.Height = Me.MultiPage1.Height - con_BorderHeightForList
    Me.ListBox2.Width = Me.MultiPage1.Width - con_BorderWidth
    Me.ListBox2.Height = Me.MultiPage1.Height - con_BorderHeightForList
End Sub

Private Sub UserForm_Deactivate()
    Set app = Nothing
End Sub

Private Sub UserForm_Resize()
    pS_Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    wAddon.Close
End Sub




Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
   MasterCheckRefresh
End Sub

