VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private t As Date
Private d As Integer
Private tx As String

Public ID As String
Public shp As Visio.Shape

Public Sub Activate(ByRef comShp As Visio.Shape, ByVal commandText As String, ByVal a_id As String)
Dim comArr() As String
    
    On Error Resume Next
    
    comArr = Split(commandText, delimiter)
    time = CDate(comArr(0))
    duration = Int(comArr(1))
    text = comArr(2)
    
    ID = a_id
    
    Set shp = comShp
End Sub

'---Свойства
Public Property Get time() As Date
On Error GoTo EX
    time = t
Exit Property
EX:
    time = 0
End Property
Public Property Let time(ByVal vNewValue As Date)
    t = vNewValue
End Property

Public Property Get duration() As Integer
On Error GoTo EX
    duration = d
Exit Property
EX:
    duration = 0
End Property
Public Property Let duration(ByVal vNewValue As Integer)
    d = vNewValue
End Property

Public Property Get text() As String
On Error GoTo EX
    text = tx
Exit Property
EX:
    text = ""
End Property
Public Property Let text(ByVal vNewValue As String)
    tx = vNewValue
End Property
