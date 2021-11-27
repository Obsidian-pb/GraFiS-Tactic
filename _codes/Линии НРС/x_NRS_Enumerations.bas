Attribute VB_Name = "x_NRS_Enumerations"
'Перечисления констант объектов железной дороги
Public Enum NRSNodeType
   nrsNone = 0                          'Обычный проежуточный узел (Рукав, разветвление и т.д.)
   nrsStarter = 1                       'Питатель (Насосы, колонки)
   nrsEnder = 2                         'Потребитель (стволы)
End Enum

'---Постоянные индеков графиков
Public Const ccs_InIdent = "Connections.GFS_In"
Public Const ccs_OutIdent = "Connections.GFS_Ou"
Public Const vb_ShapeType_Other = 0                'Ничего
Public Const vb_ShapeType_Hose = 1                 'Рукава
Public Const vb_ShapeType_PTV = 2                  'ПТВ
Public Const vb_ShapeType_Razv = 3                 'Разветвление
Public Const vb_ShapeType_Tech = 4                 'Техника
Public Const vb_ShapeType_VsasSet = 5              'Всасывающая сетка с линией
Public Const vb_ShapeType_GE = 6                   'Гидроэлеватор
Public Const vb_ShapeType_WaterContainer = 7       'Водяная емкость

'Перечисление типов данных
Public Enum NRSProp
    nrsPropS = 0                            'Сопротивление
    nrsPropQ = 1                            'Расход
    nrsPropH = 2                            'Напор внутренни
    nrsPropP = 3                            'Проводимость
    nrsPropZ = 4                            'Подъем
    nrsProphOut = 5                         'Напор на выходе
    nrsProphIn = 6                          'Напор на входе
    nrsProphLost = 7                        'Потеря напора
End Enum



'Public Event SChanged(ByVal a_S As Single)
'Public Event QChanged(ByVal a_Q As Single)
'Public Event HChanged(ByVal a_H As Single)
'Public Event PChanged(ByVal a_P As Single)
'Public Event ZChanged(ByVal a_Z As Single)
'Public Event hOutChanged(ByVal a_hOut As Single)
'Public Event hInChanged(ByVal a_hIn As Single)
'Public Event hLostChanged(ByVal a_hIn As Single)
