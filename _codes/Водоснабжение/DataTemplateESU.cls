VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTemplateESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Шаблон данных ЭСУ ППВ
'Implements IDataTemplate

Const d = ":"


Public image As String          'Изображение
Public PPVType As String
Public street As String
Public state As String
Public neisprType As String
Public lastCheckDate As String
Public prinadl As String
Public coord As String
Public tablePlace As String
Public tableForward As String
Public tableRight As String
Public tableLeft As String
Public note As String






Public Property Get Delimiter() As String
    Delimiter = d
End Property

Public Function GetDataString() As String
Dim tmpStr As String

    tmpStr = "ESU" & d & _
             image & d & _
             PPVType & d & _
             street & d & _
             state & d & _
             neisprType & d & _
             lastCheckDate & d & _
             prinadl & d & _
             coord & d & _
             tablePlace & d & _
             tableForward & d & _
             tableRight & d & _
             tableLeft & d & _
             note
             
GetDataString = tmpStr
End Function

Public Function LoadData(ByVal data As String) As Boolean
'Пытаемся олучить данные из входящей строки. Если это не получается возвращаем False
Dim dataArr() As String

    On Error GoTo EX
    
    dataArr = Split(data, d)
    
    image = dataArr(2)
    PPVType = dataArr(3)
    street = dataArr(4)
    state = dataArr(5)
    neisprType = dataArr(6)
    lastCheckDate = dataArr(7)
    prinadl = dataArr(8)
    coord = dataArr(9)
    tablePlace = dataArr(10)
    tableForward = dataArr(11)
    tableRight = dataArr(12)
    tableLeft = dataArr(13)
    note = dataArr(14)
    
    LoadData = True
Exit Function
EX:
    LoadData = False
End Function




'--------------Загрузка данных из HTML--------------------
Public Sub LoadFromHTML(ByVal html As String)
'Загружаем данные из html разметки
    image = GetPropValue(html, "<img src='", "' width")
    PPVType = GetPropValue(html, "'/><br><br>", "<br>")
    street = GetPropValue(html, "Улица (наименование объекта): ", "<br>")
    state = GetPropValue(html, "Техническое состояние: ", "<br>")
    neisprType = GetPropValue(html, "Характер неисправности: ", "<br>")
    lastCheckDate = GetPropValue(html, "Дата последней проверки: ", "<br>")
    prinadl = GetPropValue(html, "Принадлежность: ", "<br>")
    coord = GetPropValue(html, "Координаты (широта, долгота): ", "<br>")
    tablePlace = GetPropValue(html, "Место нахождения таблички: ", "<br>")
    tableForward = GetPropValue(html, "Расстояние от таблички прямо: ", "<br>")
    tableRight = GetPropValue(html, "Расстояние от таблички вправо: ", "<br>")
    tableLeft = GetPropValue(html, "Расстояние от таблички влево: ", "<br>")
    note = GetPropValue(html, "Примечание: ", Chr(10))
End Sub

Private Function ParseHTML(ByVal html As String)
Dim mainArr As String
End Function

Private Function GetPropValue(ByVal text As String, ByVal propName As String, ByVal propEnd As String) As String
Dim pos1 As Long
Dim pos2 As Long
Dim i As Long
    
    pos1 = InStr(1, text, propName) + Len(propName)
    i = pos1
    Do While i < Len(text)
        If Mid(text, i, Len(propEnd)) = propEnd Then
            pos2 = i
            GetPropValue = Mid(text, pos1, pos2 - pos1)
            Exit Function
        End If
        i = i + 1
    Loop
GetPropValue = " "
End Function
