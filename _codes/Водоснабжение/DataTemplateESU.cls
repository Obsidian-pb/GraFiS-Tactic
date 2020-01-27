VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTemplateESU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''Шаблон данных ЭСУ ППВ
''Implements IDataTemplate
'
'Const d = ":"
'
'
'Public image As String          'Изображение
'Public PPVMode As String
'Public street As String
'Public state As String
'Public neisprType As String
'Public lastCheckDate As String
'Public prinadl As String
'Public coord As String
'Public tablePlace As String
'Public tableForward As String
'Public tableRight As String
'Public tableLeft As String
'Public note As String
'Public number As String
'Public diameter As String
'Public PGType As String
'
'
'
'
'
'
'Public Property Get Delimiter() As String
'    Delimiter = d
'End Property
'
'Public Function GetDataString() As String
'Dim tmpStr As String
'
'    tmpStr = "ESU" & d & _
'             image & d & _
'             PPVMode & d & _
'             street & d & _
'             state & d & _
'             neisprType & d & _
'             lastCheckDate & d & _
'             prinadl & d & _
'             coord & d & _
'             tablePlace & d & _
'             tableForward & d & _
'             tableRight & d & _
'             tableLeft & d & _
'             note & d & _
'             number & d & _
'             diameter
'
'GetDataString = tmpStr
'End Function
'
'Public Function GetHTMLString() As String
'Dim tmpStr As String
'
'    If Len(image) > 0 Then
'        tmpStr = tmpStr & "<img src='" & image & "' width='200' style='margin: auto'/><br><br>"
'    End If
'    tmpStr = tmpStr & PPVMode & "<br/>"
'    tmpStr = tmpStr & "Улица (наименование объекта): " & street & "<br/>"
'    tmpStr = tmpStr & "Техническое состояние: " & state & "<br/>"
'    tmpStr = tmpStr & "Характер неисправности: " & neisprType & "<br/>"
'    tmpStr = tmpStr & "Дата последней проверки: " & lastCheckDate & "<br/>"
'    tmpStr = tmpStr & "Принадлежность: " & prinadl & "<br/>"
'    tmpStr = tmpStr & "Координаты (широта, долгота): " & coord & "<br/>"
'    tmpStr = tmpStr & "Место нахождения таблички: " & tablePlace & "<br/>"
'    tmpStr = tmpStr & "Расстояние от таблички прямо: " & tableForward & "<br/>"
'    tmpStr = tmpStr & "Расстояние от таблички вправо: " & tableRight & "<br/>"
'    tmpStr = tmpStr & "Расстояние от таблички влево: " & tableLeft & "<br/>"
'    tmpStr = tmpStr & "Примечание: " & note & "<br/>"
'
''    tmpStr = tmpStr & "<table width='100' height='100'  style='margin: auto; border: 3px solid red'>" & _
''              "<tr>" & _
''                "<td align='left' valign='top' width='25%'>" & if(iserr(User.PGCaption),'',User.PGCaption) & "</td>" & _
''                "<td colspan='2' align='center' valign='bottom' width='50%'> ПГ-14</td>" & _
''                "<td align='right' valign='top' width='25%'>12</td>" & _
''              "</tr>" & _
''              "<tr>" & _
''                "<td colspan='4' align='center' valign='bottom' style='border-bottom: 1px solid red'>14,5</td>" & _
''              "</tr>" & _
''              "<tr>" & _
''                "<td colspan='2' align='center' style='border-right: 1px solid red' width='50%' height='25%'></td>" & _
''                "<td colspan='2' align='center' width='50%'></td>" & _
''              "</tr>" & _
''            "</table>"
'
'
'GetHTMLString = tmpStr
'End Function
'
'Public Function LoadData(ByVal data As String) As Boolean
''Пытаемся олучить данные из входящей строки. Если это не получается возвращаем False
'Dim dataArr() As String
'
'    On Error GoTo EX
'
'    dataArr = Split(data, d)
'
'    image = dataArr(2)
'    PPVMode = dataArr(3)
'    street = dataArr(4)
'    state = dataArr(5)
'    neisprType = dataArr(6)
'    lastCheckDate = dataArr(7)
'    prinadl = dataArr(8)
'    coord = dataArr(9)
'    tablePlace = dataArr(10)
'    tableForward = dataArr(11)
'    tableRight = dataArr(12)
'    tableLeft = dataArr(13)
'    note = dataArr(14)
'    number = dataArr(14)
'    diameter = dataArr(14)
'
'    LoadData = True
'Exit Function
'EX:
'    LoadData = False
'End Function
'
'
'
'
''--------------Загрузка данных из HTML--------------------
'Public Sub LoadFromHTML(ByVal html As String)
''Загружаем данные из html разметки
'    image = GetPropValue(html, "<img src='", "' width")
'    PPVMode = GetPropValue(html, "'/><br><br>", "<br>")
'    street = GetPropValue(html, "Улица (наименование объекта): ", "<br>")
'    state = GetPropValue(html, "Техническое состояние: ", "<br>")
'    neisprType = GetPropValue(html, "Характер неисправности: ", "<br>")
'    lastCheckDate = GetPropValue(html, "Дата последней проверки: ", "<br>")
'    prinadl = GetPropValue(html, "Принадлежность: ", "<br>")
'    coord = GetPropValue(html, "Координаты (широта, долгота): ", "<br>")
'    tablePlace = GetPropValue(html, "Место нахождения таблички: ", "<br>")
'    tableForward = GetPropValue(html, "Расстояние от таблички прямо: ", "<br>")
'    tableRight = GetPropValue(html, "Расстояние от таблички вправо: ", "<br>")
'    tableLeft = GetPropValue(html, "Расстояние от таблички влево: ", "<br>")
'    note = GetPropValue(html, "Примечание: ", Chr(10))
'End Sub
'
'Private Function GetPropValue(ByVal text As String, ByVal propName As String, ByVal propEnd As String) As String
'Dim pos1 As Long
'Dim pos2 As Long
'Dim i As Long
'
'    pos1 = InStr(1, text, propName) + Len(propName)
'    i = pos1
'    Do While i < Len(text)
'        If Mid(text, i, Len(propEnd)) = propEnd Then
'            pos2 = i
'            GetPropValue = Mid(text, pos1, pos2 - pos1)
'            Exit Function
'        End If
'        i = i + 1
'    Loop
'GetPropValue = " "
'End Function
