VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireNozzleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------Модель пожарного ствола---------------------


Public position As c_Cell                   'Клетка координаты ствола
Public fireContactCell As c_Cell            'Клетка координаты контакта рабочей зоны ствола с фронтом пожара
Public nozzleType As NozzleTypes            'Тип ствола
Public waterExpense As Single               'Расход ствола л/с
Public angle As Double                      'Угол наклона ствола


