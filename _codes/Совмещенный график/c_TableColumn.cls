VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_TableColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------Класс Колонки таблицы данных для графиков-------------------------------
Public Name As String
Public index As Integer

Private TimeValue As String
Private DataValue As String

Private Const Height As Integer = 64
Private Const Width As Integer = 42

Private WithEvents TB_Time As TextBox      'Время
Attribute TB_Time.VB_VarHelpID = -1
Private WithEvents TB_Data As TextBox      'Значение
Attribute TB_Data.VB_VarHelpID = -1
Private WithEvents CB_Delete As CommandButton    'Кнопка "Удалить"
Attribute CB_Delete.VB_VarHelpID = -1









'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
'    MsgBox "created"
End Sub

Private Sub Class_Terminate()

    DataForm.Controls.Remove ("TB_Time_" & index)  '.Item("TB_Time_" & index).
    DataForm.Controls.Remove ("TB_Data_" & index)
    DataForm.Controls.Remove ("CB_Delete_" & index)
    
End Sub


'---------------------------------------Проки создания и удаления класса-------------------------------------------
Public Sub Activate(ByVal a_ind As Integer, ByVal a_Left As Integer, ByVal a_Top As Integer, _
                    ByRef a_Time As String, ByRef a_Data As String)
'Основная прока активации класса
    Name = "TableColumn_" & index
    
    index = a_ind
    TimeValue = a_Time
    DataValue = a_Data
    
    '---Создаем контролы
    '---Добавляем текстовое поле ВРЕМЯ
    Set TB_Time = DataForm.Controls.Add("Forms.TextBox.1", "TB_Time_" & index)
        TB_Time.Left = a_Left
        TB_Time.Top = a_Top
        TB_Time.Width = 42
        TB_Time.Visible = True
        TB_Time.Text = TimeValue
'        TB_Time.Text = TB_Time.Name
    
    '---Добавляем текстовое поле ЗНАЧЕНИЕ (площадь/расход)
    Set TB_Data = DataForm.Controls.Add("Forms.TextBox.1", "TB_Data_" & index)
        TB_Data.Left = a_Left
        TB_Data.Top = 24
        TB_Data.Width = 42
        TB_Data.Visible = True
        TB_Data.Text = DataValue
    
    '---Добавляем кнопку "Удалить"
    Set CB_Delete = DataForm.Controls.Add("Forms.CommandButton.1", "CB_Delete_" & index)
        CB_Delete.Left = a_Left
        CB_Delete.Top = 42
        CB_Delete.Width = 42
        CB_Delete.Height = 18
        CB_Delete.Visible = True
        CB_Delete.Caption = "-"
        
        
        
    
End Sub









'---------------------------------------Проки реагирования на события----------------------------------------------
Private Sub CB_Delete_Click()
'Реагируем на нажатеие кнопки Удалить
    DataForm.DeleteColumn index
End Sub




Private Sub TB_Data_Change()
    If IsNumeric(TB_Data) And Len(TB_Data.Text) > 0 Then
        DataForm.PS_ChangeDataValue index, TB_Data.Text
        TB_Data.ForeColor = vbBlack
    Else
        TB_Data.ForeColor = vbRed
    End If
End Sub



Private Sub TB_Time_Change()
    If IsNumeric(TB_Time.Text) And Len(TB_Time.Text) > 0 Then
        DataForm.PS_ChangeTimeValue index, TB_Time.Text
        TB_Time.ForeColor = vbBlack
    Else
        TB_Time.ForeColor = vbRed
    End If
End Sub

Private Sub TB_Time_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Переходим к соседним ячейкам времени
    Select Case KeyCode
    Case Is = 37 'left
        DataForm.PS_GoToTimeColumn index - 1
    Case Is = 38 'up
        
    Case Is = 39 'right
        DataForm.PS_GoToTimeColumn index + 1
    Case Is = 40 'down
        TB_Data.SelStart = 0
        TB_Data.SelLength = Len(TB_Data.Text)
    End Select
    
End Sub
Private Sub TB_Data_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Переходим к соседним ячейкам данных
    Select Case KeyCode
    Case Is = 37 'left
        DataForm.PS_GoToDataColumn index - 1
    Case Is = 38 'up
        TB_Time.SelStart = 0
        TB_Time.SelLength = Len(TB_Time.Text)
    Case Is = 39 'right
        DataForm.PS_GoToDataColumn index + 1
    Case Is = 40 'down
        
    End Select
End Sub















