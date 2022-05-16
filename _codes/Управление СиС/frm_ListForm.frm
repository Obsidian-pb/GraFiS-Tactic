VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ListForm 
   Caption         =   "������"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11040
   OleObjectBlob   =   "frm_ListForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ListForm"
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

Private WithEvents wAddon As Visio.Window
Attribute wAddon.VB_VarHelpID = -1

Public WithEvents menuButtonExportToWord As CommandBarButton
Attribute menuButtonExportToWord.VB_VarHelpID = -1
Public WithEvents menuButtonExportToExcel As CommandBarButton
Attribute menuButtonExportToExcel.VB_VarHelpID = -1



Private Sub Stretch()
'������������� ������ ����������� ����
    Me.LB_List.width = Me.width - con_BorderWidth
    Me.LB_List.height = Me.height - con_BorderHeightForList
End Sub


'������ �� �� �������� - �������� ����� ����� �����������. ��� �� ����� ����� �������� ��� �������...
Private Sub LB_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 17 Then ctrlOn = True
End Sub
Private Sub LB_List_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 17 Then ctrlOn = False
End Sub

Private Sub UserForm_Resize()
    Stretch
End Sub

Public Sub CloseThis()
    If wAddon Is Nothing Then Exit Sub
    wAddon.Close
End Sub


Public Function Activate(ByVal arr As Variant, ByVal colWidth As String, _
                         ByVal frmID As String, ByVal frmCaption As String) As frm_ListForm
Dim colCount As Byte
    
    
    '��������� ���������� ������
    colCount = UBound(arr, 2) + 1
    Me.LB_List.ColumnCount = colCount
    
    If colWidth <> "" Then
        Me.LB_List.ColumnWidths = colWidth
    End If
    
    Me.LB_List.List = arr
    Me.LB_List.AddItem  '��������������� ������, ��� �������������� ������� ��������� ������ � ������. (!)��� �� ������ ����� ������ ��������� ������
    
    '���������� ����������� � ���������� �����
    Set wAddon = ActiveWindow.Windows.Add(frmID, visWSVisible + visWSAnchorMerged + visWSDockedBottom, visAnchorBarAddon, , , 300, 300, "LST", "LST")
    
    Me.Caption = frmID
    FormHandle = FindWindow(vbNullString, frmID)
    SetWindowLong FormHandle, GWL_STYLE, WS_CHILD Or WS_VISIBLE
    SetParent FormHandle, wAddon.WindowHandle32
    wAddon.Caption = frmCaption
    
Me.Show
Set Activate = Me
End Function




Private Sub LB_List_Change()
Dim shpID As Long
    
    On Error GoTo ex
    
'---�������� ������
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    '---���� ����� Ctrl, �� �������� ��������� ������
    If ctrlOn Then
        Me.LB_List.MultiSelect = 1 '����� ����������� ������ �� ��������!
        Application.ActiveWindow.Select Application.ActivePage.Shapes.ItemFromID(shpID), visSelect
    Else
        Me.LB_List.MultiSelect = 0
        Application.ActiveWindow.Select Application.ActivePage.Shapes.ItemFromID(shpID), visDeselectAll + visSelect
    End If
    
ex:
End Sub

Private Sub LB_List_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim shpID As Long
Dim shp As Visio.Shape

    On Error GoTo ex
    
    '���������� ������ ��� ������� ������� ������
    shpID = Me.LB_List.Column(0, Me.LB_List.ListIndex)
    Set shp = Application.ActivePage.Shapes.ItemFromID(shpID)
    Application.ActiveWindow.Select shp, visDeselectAll + visSelect
    
    '������������� ����� �� ������
    Application.ActiveWindow.Zoom = 1.5 * GetScaleAt200
    Application.ActiveWindow.ScrollViewTo shp.Cells("PinX"), shp.Cells("PinY")
ex:
End Sub







'------------------������ � ����������� ����------------------
Private Sub LB_List_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        CreateNewMenu
    End If
End Sub


Private Sub CreateNewMenu()
'������ ����������� ���� ������� ��������
Dim popupMenuBar As CommandBar
Dim Ctrl As CommandBarControl
    
    '�������� ������ �� ����������� ����
    GetToolBar popupMenuBar, "ContextListMenu", msoBarPopup
    
    '������� ��������� ������ ����
    For Each Ctrl In popupMenuBar.Controls
        Ctrl.Delete
    Next
    
    '��������� ����� ������ (����� ���������� ��� ��������� ������� ������������ � ����� �� �����)
    Set menuButtonExportToWord = NewPopupItem(popupMenuBar, 1, 567, "�������������� � Word")        '268,751
    Set menuButtonExportToExcel = NewPopupItem(popupMenuBar, 1, 566, "�������������� � Excel")
    
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

Private Sub menuButtonExportToWord_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ExportToWord
End Sub
Private Sub menuButtonExportToExcel_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ExportToExcel
End Sub

'---------------------������� � Word---------------------------------
Public Sub ExportToWord()
'������������ ���������� ������ � �������� Word
Dim wrd As Object
Dim wrdDoc As Object
Dim wrdTbl As Object
Dim wrdTblRow As Object

Dim i As Integer
Dim j As Integer
Dim colCount As Byte

Dim s As String
    
    
    colCount = Me.LB_List.ColumnCount - 1
    
    '������� ����� �������� Word
    Set wrd = CreateObject("Word.Application")
    wrd.Visible = True
    wrd.Activate
    Set wrdDoc = wrd.Documents.Add
    wrdDoc.Activate
    '������� � ����� ��������� ������� ��������� ��������
    Set wrdTbl = wrdDoc.Tables.Add(wrd.Selection.Range, Me.LB_List.ListCount, colCount)
    With wrdTbl
        If .style <> "����� �������" Then
            .style = "����� �������"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    
    
    '��������� ������� ChrW(9500)
    i = 0
    Do Until IsNull(Me.LB_List.Column(1, i))
        For j = 1 To colCount
            s = Me.LB_List.Column(j, i)
            s = Replace(s, ChrW(9500), "")
            s = Replace(s, ChrW(9492), "")
            wrdTbl.Rows(i + 1).Cells(j).Range.text = s
        Next j
        i = i + 1
        If i > 2000 Then
            '��������� �����
            Exit Do
        End If
    Loop
    
    wrdTbl.AutoFitBehavior 1        '������������� ������ �������� �� �����������
End Sub


'---------------------������� � Excel---------------------------------
Public Sub ExportToExcel()
'������������ ���������� ������ � �������� Excel
Dim exl As Object
Dim wkb As Object
Dim sht As Object
'Dim shtRow As Object

Dim i As Integer
Dim j As Integer
Dim colCount As Byte

Dim s As String
    
    
    colCount = Me.LB_List.ColumnCount - 1
    
    '������� ����� �������� Excel
    Set exl = CreateObject("Excel.Application")
    exl.Visible = True
'    exl.ActiveWindow.WindowState = -4137
    Set wkb = exl.Workbooks.Add
    Set sht = wkb.Sheets(1)
    wkb.Activate
       
    '����������� �������
    sht.Range(sht.Cells(1, 1), sht.Cells(UBound(Me.LB_List.List, 1), colCount)).Select
    With exl.Selection.Borders()
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
    End With
    exl.Selection.NumberFormat = "@"
    
    
    '��������� ������� ChrW(9500)
    i = 0
    Do Until IsNull(Me.LB_List.Column(1, i))
        For j = 1 To colCount
            s = Me.LB_List.Column(j, i)
            s = Replace(s, ChrW(9500), "")
            s = Replace(s, ChrW(9492), "")
            sht.Cells(i + 1, j).Formula = s
        Next j
        i = i + 1
        If i > 2000 Then
            '��������� �����
            Exit Do
        End If
    Loop
    
    exl.Selection.Columns.AutoFit    '������������� ������ �������� �� �����������
End Sub



