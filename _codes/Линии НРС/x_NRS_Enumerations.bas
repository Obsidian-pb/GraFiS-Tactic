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
