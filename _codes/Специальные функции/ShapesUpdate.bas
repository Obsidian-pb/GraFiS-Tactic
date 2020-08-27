Attribute VB_Name = "ShapesUpdate"
Option Explicit
'Модуль для хранения процедур обновления фигур




Public Sub ShapesUpdate1(ShpObj As Visio.Shape)
'Прока обновляет фигуры - сейчас обновляет подписи в IndexPers
Dim vsoDoc As Visio.Document
Dim vsoShape As Visio.Shape
Dim vsoMaster As Visio.Master
Dim vsoMasterShape As Visio.Shape
    
'    On Error GoTo EX
    On Error Resume Next
    
    '---Перебираем все фигуры на листе
    For Each vsoShape In Application.ActivePage.Shapes
        '---Если фигура - входит в ГраФиС, то
        If IsGrafisShape(vsoShape) Then
            '---Перебираем все мастера в прикрепленных трафаретах
            For Each vsoDoc In Application.Documents
                For Each vsoMaster In vsoDoc.Masters
                    Set vsoMasterShape = vsoMaster.Shapes(1)
                    If IsGrafisShape(vsoMasterShape) Then
                        If vsoShape.Cells("User.IndexPers").ResultStr(visUnitsString) = _
                                vsoMasterShape.Cells("User.IndexPers").ResultStr(visUnitsString) Then
                            vsoShape.Cells("User.IndexPers.Prompt").FormulaU = _
                                vsoMasterShape.Cells("User.IndexPers.Prompt").FormulaU
                            Exit For
                            Exit For
                        End If
                    End If
                Next vsoMaster

            Next vsoDoc
            
        End If
    Next vsoShape
    
    '---Удаляем фигуру
    ShpObj.Delete
    
Exit Sub
ex:
    
End Sub





Private Function IsGrafisShape(ByRef shp As Visio.Shape) As Boolean
'Првоеряем является ли фигура - фигурой ГраФиС
    IsGrafisShape = shp.CellExists("User.IndexPers", 0) = True And shp.CellExists("User.Version", 0) = True
End Function
