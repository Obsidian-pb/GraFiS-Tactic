VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertGridForm 
   Caption         =   "Настройки решетки"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   OleObjectBlob   =   "InsertGridForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertGridForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr As Visio.Master
Private VertLabels() As String
Private HorLabels() As String
Private VertLabelIndex As Integer
Private HorLabelIndex As Integer



Private Sub B_SaveSettings_Click()
'---Сохраняем списки лэйблов для осей (Из реестра Windows)
    SaveSetting "GraFiS", "GraFiS_Section", "VertAxisLabelsString", Me.TB_VertAxisNames.Text
    SaveSetting "GraFiS", "GraFiS_Section", "HorAxisLabelsString", Me.TB_HorAxisNames.Text
    Me.L_Saved.Visible = True
End Sub

Private Sub UserForm_Activate()
'---Првоеряем подключен ли трафарет "Структурные элементы"
    If CheckStractsStencil Then
        Me.CBox_AddAxis.Enabled = True
        Me.TB_OutSpace.Enabled = True
        Me.Spin_Space.Enabled = True
        Me.Label6.Enabled = True
        Me.Label7.Enabled = True
        Me.Label8.Enabled = True
        Me.Label9.Enabled = True
        Me.Label10.Enabled = True
        Me.TB_VertAxisNames.Enabled = True
        Me.TB_HorAxisNames.Enabled = True
    Else
        Me.CBox_AddAxis.Enabled = False
        Me.TB_OutSpace.Enabled = False
        Me.Spin_Space.Enabled = False
        Me.Label6.Enabled = False
        Me.Label7.Enabled = False
        Me.Label8.Enabled = False
        Me.Label9.Enabled = False
        Me.Label10.Enabled = False
        Me.TB_VertAxisNames.Enabled = False
        Me.TB_HorAxisNames.Enabled = False
    End If
    '---Получаем списки лэйблов для осей (Из реестра Windows)
    Me.L_Saved.Visible = False
    Me.TB_VertAxisNames.Text = GetSetting("GraFiS", "GraFiS_Section", "VertAxisLabelsString", _
        "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30")
    Me.TB_HorAxisNames.Text = GetSetting("GraFiS", "GraFiS_Section", "HorAxisLabelsString", _
        "А;Б;В;Г;Д;Е;Ж;З;И;К;Л;М;Н;О;П;Р;С;Т;У;Ф;Х;Ц;Ч;Ш;Щ")
End Sub
Private Function CheckStractsStencil() As Boolean
Dim doc As Visio.Document

    For Each doc In Application.Documents
        If doc.Name = "WALL_M.VSS" Or doc.Name = "WALL_M.VSSX" Then
            CheckStractsStencil = True
            Exit Function
        End If
    Next doc

CheckStractsStencil = False
End Function
Private Sub B_Cancel_Click()
    Me.Hide
End Sub
Private Sub B_OK_Click()
    DropGrid Me.TB_X, Me.TB_Y
    Me.Hide
End Sub
Private Sub TB_HorLines_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Проверяем не заканчивается ли текст в поле на ";" и если заканчивается - удаляем его
TB_HorLines.Text = DeleteLastQuart(TB_HorLines.Text)
End Sub
Private Sub TB_VertLines_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Проверяем не заканчивается ли текст в поле на ";" и если заканчивается - удаляем его
TB_VertLines.Text = DeleteLastQuart(TB_VertLines.Text)
End Sub
Private Sub Spin_Space_SpinUp()
Me.TB_OutSpace.Text = Me.Spin_Space.Value
End Sub
Private Sub Spin_Space_SpinDown()
Me.TB_OutSpace.Text = Me.Spin_Space.Value
End Sub



Private Sub DropGrid(Optional xStart As Double = 0, Optional yStart As Double = 0)
'Основная прока вбрасывания решетки
Dim VertStrings() As String
Dim HorStrings() As String
Dim VertString As Variant
Dim HorString As Variant
Dim curPosX As Double
Dim curPosY As Double
Dim i As Integer
Dim captionDir As String
Dim caption As String
Dim axisVertColl As Collection
Dim axisHorColl As Collection
Dim tnpShape As Visio.Shape
    
    On Error GoTo EX

    '---Получаем массивы имен для осей
    FormNamesArray
    
    '---Если есть необходимость добавить оси - инициируем коллекции
    If Me.CBox_AddAxis.Value Then
        '---Получаем мастер оси
        Set mstr = Application.Documents("WALL_M.VSS").Masters("Линия сетки")
        
        '---Инициируем коллекции фигур осей
        Set axisVertColl = New Collection
        Set axisHorColl = New Collection
    End If
    
    '---Получаем массив значений для вертиальных линий (ось x)
    VertStrings = Split(DeleteLastQuart(Me.TB_VertLines.Text), ";")
    '---Получаем массив значений для горизонтальных линий (ось y)
    HorStrings = Split(DeleteLastQuart(Me.TB_HorLines.Text), ";")
    
    '---Вставляем все вертикальные линии
    curPosX = xStart
    For Each VertString In VertStrings
        'Определяем подпись
        captionDir = GetCaption(VertString)
            
        'Определяем количество повторений
        For i = 1 To GetRepeatCount(VertString)
            DropGuide True, curPosX, CStr(VertString)
            curPosX = curPosX + CDbl(VertString)
            

            If Me.CBox_AddAxis.Value Then
            '---Определяем линии осей
                If captionDir = "" Then caption = GetNextLabel(True) Else caption = captionDir
                axisVertColl.Add DropAxis(True, curPosX, caption)
            End If
        Next i
    Next VertString
    
    '---Вставляем все горизонтальные линии
    curPosY = yStart
    For Each HorString In HorStrings
        'Определяем подпись
        captionDir = GetCaption(HorString)
        
        'Определяем количество повторений
        For i = 1 To GetRepeatCount(HorString)
            DropGuide False, , , curPosY, CStr(HorString)
            curPosY = curPosY + CDbl(HorString)
            

            If Me.CBox_AddAxis.Value Then
            '---Определяем линии осей
                If captionDir = "" Then caption = GetNextLabel(False) Else caption = captionDir
                axisHorColl.Add DropAxis(False, curPosY, caption)
            End If
        Next i
    Next HorString
    
    '---Исправляем длину и положение осей сетки в коллекции
    If Me.CBox_AddAxis.Value Then
        AxisFix axisVertColl, True, yStart - CDbl(Me.TB_OutSpace.Text), curPosY + CDbl(Me.TB_OutSpace.Text)
        AxisFix axisHorColl, False, xStart - CDbl(Me.TB_OutSpace.Text), curPosX + CDbl(Me.TB_OutSpace.Text)
    End If
    
Set axisVertColl = Nothing
Set axisHorColl = Nothing
Exit Sub
EX:
    MsgBox "Невозможно построить сетку по указанным вами данным! Проверьте корректность введнных сведений", _
        vbInformation, "Предупреждение"
End Sub

'------------------------------------Проки вброса фигур--------------------------------------------------
Private Sub DropGuide(ByVal vert As Boolean, Optional xStart As Double = 0, Optional xPos As String = "0", _
                        Optional yStart As Double = 0, Optional yPos As String = "0")
'Прока вбрасывает направляющую соответственно переданных значений
Dim xPosD As Double
Dim yPosD As Double
Dim Guide As Visio.Shape

    If vert Then
    'Если вертикальные линии
        'Вбрасываем фигуру и располагаем ее на листе
        Set Guide = Application.ActivePage.AddGuide(visVert, 0, 0)
        
        xPosD = xStart + CDbl(xPos)
        Guide.Cells("PinX").FormulaU = xPosD & "mm"
    Else
    'Если горизонатльные
        'Вбрасываем фигуру и располагаем ее на листе
        Set Guide = Application.ActivePage.AddGuide(visHorz, 0, 0)
        
        xPosD = yStart + CDbl(yPos)
        Guide.Cells("PinY").FormulaU = xPosD & "mm"
    End If
    
    '---Добавляем фигуру в слой с индексом слоя "Направляющие"
    Guide.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & AddLayer("Направляющие") & """"
    
End Sub

Private Function DropAxis(ByVal vert As Boolean, ByVal mainPosition As Double, _
                        ByVal caption As String) As Visio.Shape
'Прока вброса оси сетки
Dim shp As Visio.Shape
    
    Set shp = Application.ActivePage.Drop(mstr, 0, 0)
    
    '---Указываем значения начальной и конечной точек
    If vert Then
    'Если вертикальные линии
        shp.Cells("BeginX").FormulaU = mainPosition & "mm"
        shp.Cells("BeginY").FormulaU = 0 & "mm"
        shp.Cells("EndX").FormulaU = mainPosition & "mm"
        shp.Cells("EndY").FormulaU = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) & "mm"
    Else
    'Если горизонатльные
        shp.Cells("BeginX").FormulaU = 0 & "mm"
        shp.Cells("BeginY").FormulaU = mainPosition & "mm"
        shp.Cells("EndX").FormulaU = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) & "mm"
        shp.Cells("EndY").FormulaU = mainPosition & "mm"
    End If

    shp.Cells("Prop.GridTag").FormulaU = """" & caption & """"
    
    '---Настраиваем внешний вид оси
    shp.CellsSRC(visSectionControls, 0, visCtlX).FormulaU = "-0.5 m"
    shp.Shapes(2).Cells("Char.Size").FormulaU = "Height*0.605/(ThePage!DrawingScale/ThePage!PageScale)"
    shp.Shapes(2).Cells("LineWeight").FormulaU = "0.35 pt*200/(ThePage!DrawingScale/ThePage!PageScale)"
    shp.Shapes(2).Cells("Char.Style").FormulaU = "0"
    shp.Shapes(1).Cells("LineWeight").FormulaU = "0.35 pt*200/(ThePage!DrawingScale/ThePage!PageScale)"

    '---Добавляем фигуру в слой с индексом слоя "Оси"
    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & AddLayer("Оси") & """"
    shp.Shapes(1).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & AddLayer("Оси") & """"
    shp.Shapes(2).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & AddLayer("Оси") & """"

Set DropAxis = shp
End Function


'------------------Проки корректировки фигур--------------------------------
Private Sub AxisFix(ByRef axisCollection As Collection, ByVal vert As Boolean, _
                    ByVal beginPos As Double, ByVal endPos As Double)
'Прока исправляет положение крайних точек осей в коллекции
Dim axis As Visio.Shape
    
    If vert Then
        For Each axis In axisCollection
            axis.Cells("BeginY").FormulaU = beginPos & "mm"
            axis.Cells("EndY").FormulaU = endPos & "mm"
        Next axis
    Else
        For Each axis In axisCollection
            axis.Cells("BeginX").FormulaU = beginPos & "mm"
            axis.Cells("EndX").FormulaU = endPos & "mm"
        Next axis
    End If

End Sub

'-----------------Работа со слоями-----------------------
Private Function AddLayer(ByVal layerName As String) As Integer
'Функция создает новый слой с указанным именем и возвращает его индекс
'Если такой слой уже существет, то просто возвращает его номер
Dim vsoLayer As Visio.layer

    'Проверяем наличие слоя
    For Each vsoLayer In Application.ActivePage.Layers
        If vsoLayer.Name = layerName Then
            AddLayer = vsoLayer.Index - 1
            Exit Function
        End If
    Next vsoLayer
    
    'Если нет такого слоя - создаем и возвращаем его индекс
    Set vsoLayer = Application.ActiveWindow.Page.Layers.Add(layerName)
    vsoLayer.NameU = layerName
    
AddLayer = vsoLayer.Index - 1
End Function




'---------------------Служебные проки и функции-------------------------------------
Private Function DeleteLastQuart(ByVal strIn As String) As String
'Возвращает значение строки без последней ";"
Dim tmpStr As String
    If Right(strIn, 1) = ";" Then tmpStr = Left(strIn, Len(strIn) - 1) Else tmpStr = strIn
    DeleteLastQuart = tmpStr
End Function
Private Function GetCaption(ByRef strIn As Variant) As String
'Прока возвращает подпись (для осей) и одновременно очищает значение от модификатора (C)
Dim caption As String
Dim pos1 As Integer
Dim pos2 As Integer
    
    caption = "" ' По-умолчанию = ""
    pos1 = InStr(1, strIn, "(")
    If pos1 > 0 Then
        pos2 = InStr(1, strIn, ")")
        caption = Mid(strIn, pos1 + 1, pos2 - pos1 - 1)
        strIn = Left(strIn, pos1 - 1)                          ' Очищаем строку от модификатора
    End If
GetCaption = caption
End Function
Private Function GetRepeatCount(ByRef strIn As Variant) As Integer
'Прока возвращает количество повторов вброса и одновременно очищает значение от модификатора *#
Dim count As Integer
Dim pos1 As Integer
Dim pos2 As Integer
    
    count = 1 ' По-умолчанию = 1
    pos1 = InStr(1, strIn, "*")
    If pos1 > 0 Then
        count = CInt(Right(strIn, Len(strIn) - pos1))
        strIn = Left(strIn, pos1 - 1)                          ' Очищаем строку от модификатора
    End If
GetRepeatCount = count
End Function
Private Function GetNextLabel(ByVal vert As Boolean) As String
'Функция возвращает следующее значение массива имен
    If vert Then
        GetNextLabel = VertLabels(VertLabelIndex)
        VertLabelIndex = VertLabelIndex + 1
        If VertLabelIndex >= UBound(VertLabels) Then VertLabelIndex = 0
    Else
        GetNextLabel = HorLabels(HorLabelIndex)
        HorLabelIndex = HorLabelIndex + 1
        If HorLabelIndex >= UBound(HorLabels) Then HorLabelIndex = 0
    End If
End Function
Private Sub FormNamesArray()
    '---Получаем массив строк с подписями для осей (если это необходимо)
    If Me.CBox_AddAxis.Value Then
        VertLabels = Split(DeleteLastQuart(Me.TB_VertAxisNames), ";")
        HorLabels = Split(DeleteLastQuart(Me.TB_HorAxisNames), ";")
        VertLabelIndex = 0
        HorLabelIndex = 0
    End If
End Sub
