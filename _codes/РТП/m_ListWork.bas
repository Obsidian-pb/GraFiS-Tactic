Attribute VB_Name = "m_ListWork"
Option Explicit




Public Sub ShowNewList()
Dim myArray(2, 2) As Variant

  myArray(0, 0) = "����"
  myArray(0, 1) = "������"
  myArray(0, 2) = "�������"
  myArray(1, 0) = "�����"
  myArray(1, 1) = "������"
  myArray(1, 2) = "���"
  myArray(2, 0) = "����"
  myArray(2, 1) = "����"
  myArray(2, 2) = "����"
    
Dim f As frm_ListForm

    Set f = New frm_ListForm
    
    f.Activate myArray, "0 pt"  ';200 pt;200 pt
    
End Sub
