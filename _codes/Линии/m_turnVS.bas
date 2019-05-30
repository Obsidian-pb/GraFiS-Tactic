Attribute VB_Name = "m_turnVS"
'---процедура автоматического изменения направления потока ВС
Public Sub turnVS(ByRef aO_Conn As Visio.Connect)

  If aO_Conn.ToSheet.CellExists("User.IndexPers", 0) = False Then Exit Sub
  
'---скрытие стрелки направления потока при подключении вставки
  If aO_Conn.ToSheet.Cells("User.IndexPers").Result(visNumber) = 191 Then aO_Conn.ToSheet.Cells("Actions.DirectionShow.Checked").Formula = 0
   
 ' If aO_Conn.ToSheet.Cells("User.IndexPers").Result(visNumber) = 105 Then
'---скрытие стрелки направления потока при подключении ВС
 '    aO_Conn.ToSheet.Cells("Actions.DirectionShow.Checked").Formula = 0

'  If aO_Conn.ToSheet.Cells("Scratch.A1").Result(visNumber) = 1 Then
'  Debug.Print aO_Conn.ToSheet
'  Debug.Print "1=" & aO_Conn.ToSheet.Cells("Scratch.A1").Result(visNumber)
'  Debug.Print "2=" & aO_Conn.ToSheet.Cells("Scratch.A2").Result(visNumber)
'  Debug.Print "3=" & aO_Conn.ToSheet.Cells("Scratch.A3").Result(visNumber)
'   If aO_Conn.ToSheet.Cells("Scratch.A2").Result(visNumber) = 1 Or _
'      aO_Conn.ToSheet.Cells("Scratch.A3").Result(visNumber) = 1 Then
'     Debug.Print 1
'      If aO_Conn.ToSheet.Cells("Scratch.C1").Result(visNumber) = 0 And _
'       (aO_Conn.ToSheet.Cells("Scratch.C2").Result(visNumber) = 0 And _
'        aO_Conn.ToSheet.Cells("Scratch.C3").Result(visNumber) = 0) Then
'        Debug.Print 2
'      End If
'     End If
'  End If
' End If
End Sub

