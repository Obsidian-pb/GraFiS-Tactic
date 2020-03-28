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

'---------------------------------------- ласс дл€ объекта настроек класса Infocollector----------------
Public GDZSRezRoundUp As Boolean             'ќкругление резервных звеньев в большую сторону (True - дл€  каждого из блоков по три звена выставл€етс€ резерв, False - резерв выставл€етс€ при наличии очередных трех звеньев)



Private Sub Class_Initialize()
    '«агружаем стартовые свойства из реестар Windows
    GDZSRezRoundUp = GetSetting("GraFiS", "GFS_CalcOptions", "GDZSRezRoundUp", 1)
End Sub

Private Sub Class_Terminate()
    '—охран€ем свойства в реестре Windows
    SaveSetting "GraFiS", "GFS_CalcOptions", "GDZSRezRoundUp", GDZSRezRoundUp
End Sub


'ѕуть в реестре: \HKEY_CURRENT_USER\Software\VB and VBA Program Settings\GraFiS\GFS_CalcOptions
