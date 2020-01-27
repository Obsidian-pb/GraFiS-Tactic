VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_CommonDataRedactESU 
   Caption         =   "Расширенные свойства"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   OleObjectBlob   =   "f_CommonDataRedactESU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_CommonDataRedactESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Шаблон представления данных ЭСУ ППВ
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
''Процедура отображения формы
'Dim dt As IDataTemplate
'
'    'Создаем экземпляр объекта шаблона данных для загрузки данных
'    Set dt = New DataTemplateESU
'    dt.LoadData data
'
'    'Сохраняем полученный объект данных как текущий объект данных формы с шаблоном DataTemplateESU
'    Set dataTemplate = dt
'
'    Me.Show
'End Sub

Public Sub FormShow(ByRef shp As Visio.Shape)
'Процедура отображения формы
'Dim dt As IDataTemplate
Dim str As String

    'Сохраняем ссылку на целевую фигуру
    Set targetShape = shp

    'Создаем шаблон данных
    Set dataTemplate = New DataTemplateESU

    'Пытаемся получить данные
    If shp.CellExists("Prop.Common", 0) Then
        If shp.CellExists("User.INPPWData", 0) Then
            'Проверяем имеются ли в фигуре сохраненные
            If shp.Cells("User.INPPWData").ResultStr(visUnitsString) > "" Then
                'Загружаем данные из текстовой строки данных
                dataTemplate.LoadData shp.Cells("User.INPPWData").ResultStr(visUnitsString)
            Else
                'Загружаем данные из html разметки
                dataTemplate.LoadFromHTML shp.Cells("Prop.Common").ResultStr(visUnitsString)
                'Дополняем их из сведений о фигуре
'                dataTemplate
'                dataTemplate.number = shp.Cells("Prop.PGNumber").ResultStr(visUnitsString)
'                dataTemplate.diameter = shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString)
            End If
        End If
    End If

    'Передаем данные из шаблона в элементы управления формы
    Me.tb_Coord = dataTemplate.coord
    Me.tb_Image = dataTemplate.image
    Me.tb_LastCheckDate = dataTemplate.lastCheckDate
    Me.tb_Neispr = dataTemplate.neisprType
    Me.tb_Note = dataTemplate.note
    Me.tb_PPVMode = dataTemplate.PPVMode
    Me.tb_Prinadl = dataTemplate.prinadl
    Me.tb_State = dataTemplate.state
    Me.tb_Street = dataTemplate.street
    Me.tb_TableForward = dataTemplate.tableForward
    Me.tb_TableLeft = dataTemplate.tableLeft
    Me.tb_TablePlace = dataTemplate.tablePlace
    Me.tb_TableRight = dataTemplate.tableRight


    'Показываем форму
    Me.Show
End Sub





Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    'Передаем данные из формы в шаблон
    dataTemplate.coord = Me.tb_Coord
    dataTemplate.image = Me.tb_Image
    dataTemplate.lastCheckDate = Me.tb_LastCheckDate
    dataTemplate.neisprType = Me.tb_Neispr
    dataTemplate.note = Me.tb_Note
    dataTemplate.PPVMode = Me.tb_PPVMode
    dataTemplate.prinadl = Me.tb_Prinadl
    dataTemplate.state = Me.tb_State
    dataTemplate.street = Me.tb_Street
    dataTemplate.tableForward = Me.tb_TableForward
    dataTemplate.tableLeft = Me.tb_TableLeft
    dataTemplate.tablePlace = Me.tb_TablePlace
    dataTemplate.tableRight = Me.tb_TableRight

    'Сохраняем данные из шаблона в фигуру
    targetShape.Cells("User.INPPWData").Formula = dataTemplate.GetDataString
    targetShape.Cells("Prop.Common").Formula = ""

    'Закрываем форму
    Me.Hide
End Sub


'Private Function GetPGType() As String
'Dim t As String
'
'    On Error GoTo EX
'
'    t = targetShape.Cells("Prop.PGType").ResultStr(visUnitsString)
'    t = Left(t, 1)
'
'    If Not (t = "М" Or t = "Н" Or t = "Л" Or t = "С") Then
'        t = "ПГ"
'    End If
'
'    GetPGType = t
'Exit Function
'EX:
'    GetPGType = ""
'End Function

