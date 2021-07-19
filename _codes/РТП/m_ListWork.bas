Attribute VB_Name = "m_ListWork"
Option Explicit




Public Sub ShowNewList()
Dim myArray(2, 2) As Variant

  myArray(0, 0) = "Зима"
  myArray(0, 1) = "Январь"
  myArray(0, 2) = "Февраль"
  myArray(1, 0) = "Весна"
  myArray(1, 1) = "Апрель"
  myArray(1, 2) = "Май"
  myArray(2, 0) = "Лето"
  myArray(2, 1) = "Июнь"
  myArray(2, 2) = "Июль"
    
Dim f As frm_ListForm

    Set f = New frm_ListForm
    
    f.Activate myArray, "0 pt"  ';200 pt;200 pt
    
End Sub
