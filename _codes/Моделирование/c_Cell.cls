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

Public X As Long
Public Y As Long




Public Property Set ParentMatrix(ByRef matrixObject As c_Matrix)
'Свойство присвоения ссылки на матрицу
    Set matrix = matrixObject
End Property

Public Property Get IsCellMaxFire() As Boolean
    IsCellMaxFire = matrix.IsCellMaxFire(X, Y)
End Property

Public Function NCell(ByVal Dir As Directions) As c_Cell
Dim tmpCell As c_Cell
Dim dx As Long
Dim dy As Long
    
    If Dir = s Then Set NCell = Me
    
    Select Case Dir
        Case Is = l
            dx = X - 1
            dy = Y
        Case Is = lu
            dy = Y + 1
            dx = X - 1
        Case Is = u
            dx = X
            dy = Y + 1
        Case Is = ru
            dy = Y + 1
            dx = X + 1
        Case Is = r
            dx = X + 1
            dy = Y
        Case Is = rd
            dy = Y - 1
            dx = X + 1
        Case Is = d
            dx = X
            dy = Y - 1
        Case Is = ld
            dy = Y - 1
            dx = X - 1
        Case Else

    End Select
    
    Set tmpCell = New c_Cell
    tmpCell.X = dx
    tmpCell.Y = dy
    Set tmpCell.ParentMatrix = matrix
    Set NCell = tmpCell
    
End Function


