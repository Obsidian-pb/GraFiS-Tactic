Attribute VB_Name = "m_CalculatorsTools"
Option Explicit





Public Sub TastCalculator()
Dim newAnalizer As ElementsShell
Dim shp As Visio.Shape
Dim shpIndex As Integer
    
    Set newAnalizer = New ElementsShell
    
    
    For Each shp In Application.ActivePage.Shapes
        If IsGFSShape(shp) Then
            shpIndex = shp.Cells("User.IndexPers")   '���������� ������ ������ ������
            
            With newAnalizer
            Select Case shpIndex
            '---�������� ����������-----------
                Case Is = ipAC   '������������
'                    .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount"
'                    .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                    .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                    .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
            End Select
            .RaiseByCellID shp, "User.FireSquare;User.PodOut;Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
            End With
        End If
    Next shp
    
'    Debug.Print "����� ��:" & newAnalizer.ByID("MainPAHave").Result
'    Debug.Print "��: " & newAnalizer.ByID("ACCount").Result
'    Debug.Print "����� �������: " & newAnalizer.ByID("TotalTechCount").Result
'    Debug.Print "����� �������� �������: " & newAnalizer.ByID("TotalFireCount").Result
'    Debug.Print "������� ������� �������: " & newAnalizer.ByID("PersonnelHave").Result
'    Debug.Print "���� �������: " & newAnalizer.ByID("WaterValueHave").Result
'
'    Debug.Print "������� 38��: " & newAnalizer.ByCellID("Prop.Hose38").Result
'    Debug.Print "������� 51��: " & newAnalizer.ByCellID("Prop.Hose51").Result
'    Debug.Print "������� 66��: " & newAnalizer.ByCellID("Prop.Hose66").Result
'    Debug.Print "������� 77��: " & newAnalizer.ByCellID("Prop.Hose77").Result
    
'    Debug.Print "������� ������: " & newAnalizer.ByID("FireSquare").Result
    Debug.Print "��������� ������: " & newAnalizer.ByID("FactStreamW").Result
    
'!    Debug.Print newAnalizer.CallNames(";")
'    Debug.Print newAnalizer.ByID("FireSquare").Result
'    Debug.Print newAnalizer.GetCalculatorByID("gdzs").Result & "=>" & newAnalizer.GetCalculatorByID("gdzs").callName
End Sub



Public Sub TastInfoCollector()

    A.Refresh(1).PrintState
'    Debug.Print A.Result("FactStreamW")

End Sub


