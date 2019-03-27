Attribute VB_Name = "m_Common"
Option Explicit


Public Sub ShowPressureForm(ShpObj As Visio.Shape)
'    f_PressureChange.SetShp (ShpObj)
    Set f_PressureChange.currentShp = ShpObj
    f_PressureChange.Show
End Sub


