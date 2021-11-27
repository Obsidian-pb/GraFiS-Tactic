Attribute VB_Name = "x_NRS_Enumerations"
'������������ �������� �������� �������� ������
Public Enum NRSNodeType
   nrsNone = 0                          '������� ������������ ���� (�����, ������������ � �.�.)
   nrsStarter = 1                       '�������� (������, �������)
   nrsEnder = 2                         '����������� (������)
End Enum

'---���������� ������� ��������
Public Const ccs_InIdent = "Connections.GFS_In"
Public Const ccs_OutIdent = "Connections.GFS_Ou"
Public Const vb_ShapeType_Other = 0                '������
Public Const vb_ShapeType_Hose = 1                 '������
Public Const vb_ShapeType_PTV = 2                  '���
Public Const vb_ShapeType_Razv = 3                 '������������
Public Const vb_ShapeType_Tech = 4                 '�������
Public Const vb_ShapeType_VsasSet = 5              '����������� ����� � ������
Public Const vb_ShapeType_GE = 6                   '�������������
Public Const vb_ShapeType_WaterContainer = 7       '������� �������

'������������ ����� ������
Public Enum NRSProp
    nrsPropS = 0                            '�������������
    nrsPropQ = 1                            '������
    nrsPropH = 2                            '����� ���������
    nrsPropP = 3                            '������������
    nrsPropZ = 4                            '������
    nrsProphOut = 5                         '����� �� ������
    nrsProphIn = 6                          '����� �� �����
    nrsProphLost = 7                        '������ ������
End Enum



'Public Event SChanged(ByVal a_S As Single)
'Public Event QChanged(ByVal a_Q As Single)
'Public Event HChanged(ByVal a_H As Single)
'Public Event PChanged(ByVal a_P As Single)
'Public Event ZChanged(ByVal a_Z As Single)
'Public Event hOutChanged(ByVal a_hOut As Single)
'Public Event hInChanged(ByVal a_hIn As Single)
'Public Event hLostChanged(ByVal a_hIn As Single)
