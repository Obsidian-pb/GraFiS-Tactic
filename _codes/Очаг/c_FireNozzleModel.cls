VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireNozzleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------������ ��������� ������---------------------


Public position As c_Cell                   '������ ���������� ������
Public fireContactCell As c_Cell            '������ ���������� �������� ������� ���� ������ � ������� ������
Public nozzleType As NozzleTypes            '��� ������
Public waterExpense As Single               '������ ������ �/�
Public angle As Double                      '���� ������� ������
Public timeOfSet As Single                  '����� ������ ������ � ���� �+ (���)

