VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoCollectorOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------����� ��� ������� �������� ������ Infocollector----------------
Public GDZSRezRoundUp As Boolean             '���������� ��������� ������� � ������� ������� (True - ���  ������� �� ������ �� ��� ����� ������������ ������, False - ������ ������������ ��� ������� ��������� ���� �������)



Private Sub Class_Initialize()
    '��������� ��������� �������� �� ������� Windows
    GDZSRezRoundUp = GetSetting("GraFiS", "GFS_CalcOptions", "GDZSRezRoundUp", 1)
End Sub

Private Sub Class_Terminate()
    '��������� �������� � ������� Windows
    SaveSetting "GraFiS", "GFS_CalcOptions", "GDZSRezRoundUp", GDZSRezRoundUp
End Sub


'���� � �������: \HKEY_CURRENT_USER\Software\VB and VBA Program Settings\GraFiS\GFS_CalcOptions
