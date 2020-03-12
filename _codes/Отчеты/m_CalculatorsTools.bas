Attribute VB_Name = "m_CalculatorsTools"
Option Explicit





Public Sub TastCalculator()
Dim newAnalizer As ElementsShell
Dim shp As Visio.Shape
    
    Set newAnalizer = New ElementsShell
    
    
'    For Each shp In Application.ActivePage.Shapes
'        newAnalizer.Calculate shp
'    Next shp
    
    Debug.Print newAnalizer.CallNames(";")
    Debug.Print newAnalizer.ByID("FireSquare").Result
'    Debug.Print newAnalizer.GetCalculatorByID("gdzs").Result & "=>" & newAnalizer.GetCalculatorByID("gdzs").callName
End Sub



