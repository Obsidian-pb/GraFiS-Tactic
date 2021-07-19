VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_SimpleDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const callName_text = " - "

Private t As Date
Private d As Integer
Private tx As String
Private sd_Type As Byte        '1 - Command, 0 - Info

Public ID As String
Public shp As Visio.Shape

Public Sub Activate(ByRef comShp As Visio.Shape, ByVal commandText As String, ByVal a_id As String)
Dim comArr() As String
Dim callName As String
    
    On Error Resume Next
    
    comArr = Split(commandText, delimiter)
    time = CDate(comArr(0))
    duration = Int(comArr(1))
    callName = getCallName(comShp)
    If callName = "-" Then
        text = comArr(2)
    Else
        text = getCallName(comShp) & callName_text & comArr(2)
    End If
    
    sd_Type = 1     'Command
    
    ID = a_id
        
    Set shp = comShp
End Sub

Public Sub ActivateAsInfo(ByRef comShp As Visio.Shape, ByVal infoText As String, ByVal a_id As String)
Dim comArr() As String
Dim callName As String
    
    On Error Resume Next
    
    comArr = Split(infoText, delimiter)
    time = CDate(comArr(0))
    callName = getCallName(comShp)
'    If callName = "-" Then
        text = comArr(1)
'    Else
'        text = getCallName(comShp) & callName_text & comArr(1)
'    End If
    
    sd_Type = 0     'Info
    
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

Public Property Get sdType() As Byte
On Error GoTo EX
    sdType = sd_Type
Exit Property
EX:
    sdType = 1
End Property

