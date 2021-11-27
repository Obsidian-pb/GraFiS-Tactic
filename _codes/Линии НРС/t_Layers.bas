Attribute VB_Name = "t_Layers"
'--------------------------------Работа со слоями-------------------------------------
Private Function GetLayerNamber(ByVal layerName As String) As Integer
'Return layer with name=layerName number in sheet layers list
Dim layer As Visio.layer
Dim i As Integer
    
    'Search layer with need name and return its number (not index because sometime it leads to mistakes)
    For i = 1 To Application.ActivePage.Layers.Count
        If Application.ActivePage.Layers(i).Name = layerName Then
            GetLayerNamber = i - 1
            Exit Function
        End If
    Next i
    
    'If layer is not found then create new one
    Set layer = Application.ActivePage.Layers.Add(layerName)
    GetLayerNumber = layer.Index - 1
End Function

Public Sub ClearLayer(ByVal layerName As String, Optional ByVal delFlag As Boolean = False)
'Delete all layer shapes
'If delFlag = True, then layer deleted too
Dim vsoSelection As Visio.Selection

    On Error Resume Next

    If delFlag Then
        Application.ActivePage.Layers.item(layerName).Delete (1)
    Else
        Set vsoSelection = Application.ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, layerName)
        vsoSelection.Delete
    End If
End Sub
