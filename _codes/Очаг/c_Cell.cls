VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private matrix As c_Matrix

Public x As Long
Public y As Long




Public Property Set ParentMatrix(ByRef matrixObject As c_Matrix)
'Свойство присвоения ссылки на матрицу
    Set matrix = matrixObject
End Property





