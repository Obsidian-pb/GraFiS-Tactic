Attribute VB_Name = "archiv"
Public Sub TTT()
Dim con As Connect
Dim shp As Visio.Shape
Dim cll As Visio.cell

    Set shp = Application.ActiveWindow.Selection(1)

'    For Each con In shp.Connects
''        If (Not NRSModel_.InModel(con.ToSheet)) And IsGFSShape(con.ToSheet) Then
'            Debug.Print con.FromCell.Name
''            NRSModel_.AddNRSNode con.ToSheet ���������
''            GetTechShapeForGESystem con.ToSheet
''        End If
'    Next con
'    For Each con In shp.FromConnects
''        If (Not NRSModel_.InModel(con.FromSheet)) And IsGFSShape(con.ToSheet) Then
'            Debug.Print con.ToCell.Name
''            NRSModel_.AddNRSNode con.FromSheet ���������
''            GetTechShapeForGESystem con.FromSheet
''        End If
'    Next con
    
    
    If IsGFSShapeWithIP(shp, indexPers.ipRukavLineNapor) Then       '���� ������� ������ �����, �� ��������� ����������� � ��� ������...
        For Each con In shp.Connects
'            If con.ToCell.Name = cll.Name Then
                Debug.Print con.FromCell.Name
'            End If
        Next con
    Else                                                            '...����� ��������� ����������� � ������ ����������
        connPointsCount = shp.RowCount(visSectionConnectionPts)
        If connPointsCount > 0 Then
            For i = 0 To connPointsCount - 1
                Set cll = shp.CellsSRC(visSectionConnectionPts, i, 0)
                If Left(cll.Name, 18) = "Connections.GFS_Ou" Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
                            Debug.Print cll.Name
                        End If
                    Next con
                ElseIf Left(cll.Name, 18) = "Connections.GFS_In" Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
                            Debug.Print cll.Name
                        End If
                    Next con
                End If
            Next i
        End If

        
        
    End If
End Sub

Public Sub SSS()
Dim NRSDemon As c_NRSDemon
Dim shp As Visio.Shape

    Set shp = Application.ActiveWindow.Selection(1)
    
    Set NRSDemon = New c_NRSDemon
    '������ ������
    NRSDemon.BuildNRSModel shp
    '������������ ������
    NRSDemon.CalculateNRSModel
    
    Set NRSDemon = Nothing
End Sub
