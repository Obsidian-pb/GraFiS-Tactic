Attribute VB_Name = "m_CalculatorsTools"
Option Explicit





Public Sub TastCalculator()
Dim newAnalizer As Calculators
Dim shp As Visio.Shape
    
    Set newAnalizer = New Calculators
    
    For Each shp In Application.ActivePage.Shapes
        newAnalizer.Calculate shp
    Next shp
    
'    Debug.Print newAnalizer.CallNamesArray(";")
    Debug.Print newAnalizer.GetCalculatorByID("FireSquare").Result
'    Debug.Print newAnalizer.GetCalculatorByID("gdzs").Result & "=>" & newAnalizer.GetCalculatorByID("gdzs").callName
End Sub
