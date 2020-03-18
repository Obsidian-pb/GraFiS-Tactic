Attribute VB_Name = "m_CalculatorsTools"
Option Explicit





Public Sub TastCalculator()
Dim newAnalizer As ElementsShell
Dim shp As Visio.Shape
Dim shpIndex As Integer
    
    Set newAnalizer = New ElementsShell
    
    
    For Each shp In Application.ActivePage.Shapes
        If IsGFSShape(shp) Then
            shpIndex = shp.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
            
            With newAnalizer
            Select Case shpIndex
            '---Пожарные автомобили-----------
                Case Is = ipAC   'Автоцистерны
'                    .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount"
'                    .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                    .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                    .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
            End Select
            .RaiseByCellID shp, "User.FireSquare;User.PodOut;Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
            End With
        End If
    Next shp
    
'    Debug.Print "Всего ПА:" & newAnalizer.ByID("MainPAHave").Result
'    Debug.Print "АЦ: " & newAnalizer.ByID("ACCount").Result
'    Debug.Print "Всего техники: " & newAnalizer.ByID("TotalTechCount").Result
'    Debug.Print "Всего пожарной техники: " & newAnalizer.ByID("TotalFireCount").Result
'    Debug.Print "Личного состава имеется: " & newAnalizer.ByID("PersonnelHave").Result
'    Debug.Print "Воды имеется: " & newAnalizer.ByID("WaterValueHave").Result
'
'    Debug.Print "Рукавов 38мм: " & newAnalizer.ByCellID("Prop.Hose38").Result
'    Debug.Print "Рукавов 51мм: " & newAnalizer.ByCellID("Prop.Hose51").Result
'    Debug.Print "Рукавов 66мм: " & newAnalizer.ByCellID("Prop.Hose66").Result
'    Debug.Print "Рукавов 77мм: " & newAnalizer.ByCellID("Prop.Hose77").Result
    
'    Debug.Print "площадь пожара: " & newAnalizer.ByID("FireSquare").Result
    Debug.Print "требуемый расход: " & newAnalizer.ByID("FactStreamW").Result
    
'!    Debug.Print newAnalizer.CallNames(";")
'    Debug.Print newAnalizer.ByID("FireSquare").Result
'    Debug.Print newAnalizer.GetCalculatorByID("gdzs").Result & "=>" & newAnalizer.GetCalculatorByID("gdzs").callName
End Sub



Public Sub TastInfoCollector()

    A.Refresh(1).PrintState
'    Debug.Print A.Result("FactStreamW")

End Sub


