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
            End If
        End If
    End If
    
    'Передаем данные из шаблона в элементы управления формы
    
    
    'Показываем форму
    Me.Show
End Sub





