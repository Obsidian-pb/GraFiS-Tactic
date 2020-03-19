Attribute VB_Name = "m__Tests"
Option Explicit





Public Sub TastCalculator()
Dim newAnalizer As ElementsShell
Dim shp As Visio.Shape
Dim shpIndex As Integer
    
    Set newAnalizer = New ElementsShell
    
'    Debug.Print newAnalizer.CallNames(";")
    
    Dim col As Collection
    Set col = newAnalizer.GetElementsCollection("")
'    Debug.Print col.Count
    
    Dim elem As Element
    Dim i As Integer
    For Each elem In col
        i = i + 1
        Debug.Print i
        elem.PrintState
    Next elem
'    For Each shp In Application.ActivePage.Shapes
'        If IsGFSShape(shp) Then
'            shpIndex = shp.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
'
'            With newAnalizer
'            Select Case shpIndex
'            '---Пожарные автомобили-----------
'                Case Is = ipAC   'Автоцистерны
''                    .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount"
''                    .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
''                    .Raise "WaterValueHave", shp.Cells("Prop.Water")
''                    .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
'            End Select
'            .RaiseByCellID shp, "User.FireSquare;User.PodOut;Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
'            End With
'        End If
'    Next shp
    
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
'    Debug.Print "требуемый расход: " & newAnalizer.ByID("FactStreamW").Result
    
'!    Debug.Print newAnalizer.CallNames(";")
'    Debug.Print newAnalizer.ByID("FireSquare").Result
'    Debug.Print newAnalizer.GetCalculatorByID("gdzs").Result & "=>" & newAnalizer.GetCalculatorByID("gdzs").callName
End Sub



Public Sub TestInfoCollector()
    A.Refresh (1)
'    A.Refresh(1).PrintState
'    Debug.Print A.Result("FactStreamW", True)

'    A.PrintState "PersonnelNeed;GDZSChainsCountWork;GDZSChainsCountNeed;GDZSChainsRezCountHave;GDZSChainsRezCountNeed;GDZSMansCountWork;GDZSMansCountNeed;GDZSMansRezCountHave;GDZSMansRezCountNeed"
    
    A.PrintState ("HosesHave;PersonnelNeed;PAHighHave;TechnicsNotMchsHave;TechnicsNotMchsOtherHave;ACNeed;ANRNeed;StvolWBNeed;StvolWANeed;StvolWLNeed;PANeedOnWaterSource")
'    Debug.Print A.Refresh(1).Result("Площадь пожара")
    
    
    
End Sub

'        .Raise "PersonnelNeed", .Result("GDZSMansRezCountNeed")
'        .SetVal "PAHighHave", .Sum("ALCount;AKPCount")
'        .SetVal "TechnicsNotMchsHave", .Result("TotalTechCount") - .Result("TotalFireCount")
'        .SetVal "TechnicsNotMchsOtherHave", .Result("TechnicsNotMchsHave") - .Result("MVDCount") - .Result("SMPCount")
'        .SetVal "ACNeed", PF_RoundUp(.Result("PersonnelNeed") / 4)
'        .SetVal "ANRNeed", PF_RoundUp(.Result("PersonnelNeed") / 5)
'        .SetVal "StvolWBNeed", PF_RoundUp(.Result("NeedStreamW") / 3.7)
'        .SetVal "StvolWANeed", PF_RoundUp(.Result("NeedStreamW") / 7.4)
'        .SetVal "StvolWLNeed", PF_RoundUp(.Result("NeedStreamW") / 12)
'        .SetVal "PANeedOnWaterSource", PF_RoundUp(.Result("NeedStreamW") / 32)
