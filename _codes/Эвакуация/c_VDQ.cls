VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_VDQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------Класс для получения значений Интенсивности (q) и Скорости (V) в зависиомсти от Плотности (D)-----------

Private Ds As Variant
'Горизонтальный путь
Private V0 As Variant
Private q0 As Variant
'Дверной проем
Private q1 As Variant
'Лестница вниз
Private V2 As Variant
Private q2 As Variant
'Лестница вверх
Private V3 As Variant
Private q3 As Variant






Private Sub Class_Initialize()
    Ds = Array(0.01, 0.05, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9)
    
    V0 = Array(100, 100, 80, 60, 47, 40, 33, 28, 23, 19, 15)
    q0 = Array(1, 5, 8, 12, 14.1, 16, 16.5, 16.3, 16.1, 15.2, 13.5)
    
    q1 = Array(1, 5, 8.7, 13.4, 16.5, 18.4, 19.6, 19.05, 18.5, 17.3, 8.5)
    
    V2 = Array(100, 100, 95, 68, 52, 40, 31, 24.5, 18, 13, 8)
    q2 = Array(1, 5, 9.5, 13.6, 15.6, 16, 15.6, 14.1, 12.6, 10.4, 7.2)
    
    V3 = Array(60, 60, 53, 40, 32, 26, 22, 18.5, 15, 13, 11)
    q3 = Array(0.6, 3, 5.3, 8, 9.6, 10.4, 11, 10.75, 10.5, 10.4, 9.9)
    
'    Debug.Print UBound(Ds)
'    Debug.Print UBound(V0)
'    Debug.Print UBound(q0)
'    Debug.Print UBound(q1)
'    Debug.Print UBound(V2)
'    Debug.Print UBound(q2)
'    Debug.Print UBound(V3)
'    Debug.Print UBound(q3)
End Sub

Public Function GetV(ByVal DD As Single, ByVal wayClass As wayClasses)
Dim curD As Single
Dim i As Integer
Dim j As Integer
    
    If DD = 0 Then
        GetV = 0
        Exit Function
    ElseIf Ds(0) > DD Then
        Select Case wayClass
            Case Is = wayClasses.hor
                GetV = V0(0)
            Case Is = wayClasses.door
                GetV = 0
            Case Is = wayClasses.laddo
                GetV = V2(0)
            Case Is = wayClasses.ladup
                GetV = V3(0)
        End Select
        Exit Function
    ElseIf Ds(UBound(Ds)) < DD Then
        Select Case wayClass
            Case Is = wayClasses.hor
                GetV = V0(UBound(Ds))
            Case Is = wayClasses.door
                GetV = 0
            Case Is = wayClasses.laddo
                GetV = V2(UBound(Ds))
            Case Is = wayClasses.ladup
                GetV = V3(UBound(Ds))
        End Select
        Exit Function
    Else
        For i = 0 To UBound(Ds)
            If Ds(i) > DD Then
                i = i - 1
                Exit For
            End If
        Next i
    End If
    j = i + 1
    
    
    Select Case wayClass
        Case Is = wayClasses.hor
            GetV = Interpolate(DD, Ds(i), Ds(j), V0(i), V0(j))
        Case Is = wayClasses.door
            GetV = 0   'Interpolate(DD, Ds(i), Ds(j), V0(i), V0(j))
        Case Is = wayClasses.laddo
            GetV = Interpolate(DD, Ds(i), Ds(j), V2(i), V2(j))
        Case Is = wayClasses.ladup
            GetV = Interpolate(DD, Ds(i), Ds(j), V3(i), V3(j))
    End Select
    
    
End Function


Public Function GetQ(ByVal DD As Single, ByVal wayClass As wayClasses)
Dim curD As Single
Dim i As Integer
Dim j As Integer
    
    If DD = 0 Then
        GetQ = 0
        Exit Function
    ElseIf Ds(0) > DD Then
        Select Case wayClass
            Case Is = wayClasses.hor
                GetQ = q0(0)
            Case Is = wayClasses.door
                GetQ = q1(0)
            Case Is = wayClasses.laddo
                GetQ = q2(0)
            Case Is = wayClasses.ladup
                GetQ = q3(0)
        End Select
        Exit Function
    ElseIf Ds(UBound(Ds)) < DD Then
        Select Case wayClass
            Case Is = wayClasses.hor
                GetQ = q0(UBound(Ds))
            Case Is = wayClasses.door
                GetQ = q1(UBound(Ds))
            Case Is = wayClasses.laddo
                GetQ = q2(UBound(Ds))
            Case Is = wayClasses.ladup
                GetQ = q3(UBound(Ds))
        End Select
        Exit Function
    Else
        For i = 0 To UBound(Ds)
            If Ds(i) > DD Then
                i = i - 1
                Exit For
            End If
        Next i
    End If
    
    If i >= UBound(Ds) Then
        'Избегаем переполнения в случае, если DDслишком велико
        i = UBound(Ds)
        Select Case wayClass
            Case Is = wayClasses.hor
                GetQ = q0(i)
            Case Is = wayClasses.door
                GetQ = q1(i)
            Case Is = wayClasses.laddo
                GetQ = q2(i)
            Case Is = wayClasses.ladup
                GetQ = q3(i)
        End Select
    Else
        j = i + 1
    End If
    
    Select Case wayClass
        Case Is = wayClasses.hor
            GetQ = Interpolate(DD, Ds(i), Ds(j), q0(i), q0(j))
        Case Is = wayClasses.door
            GetQ = Interpolate(DD, Ds(i), Ds(j), q1(i), q1(j))
        Case Is = wayClasses.laddo
            GetQ = Interpolate(DD, Ds(i), Ds(j), q2(i), q2(j))
        Case Is = wayClasses.ladup
            GetQ = Interpolate(DD, Ds(i), Ds(j), q3(i), q3(j))
    End Select
    
    
End Function


Public Function GetVByQ(ByVal QQ As Single, ByVal wayClass As wayClasses)
'Ищем скорость(V) по интенсивности(q)
Dim curQprev As Single
Dim curQ As Single
Dim curVprev As Single
Dim curV As Single
Dim i As Integer
    
    
    
    If QQ = 0 Then
        Debug.Print "!!!q = 0"
        GetVByQ = 0
        Exit Function
    End If
    
    
    Select Case wayClass
        Case Is = wayClasses.hor
            'Ищем значение q0 больше переданного значения интенсивности
'            For i = 0 To UBound(Ds)
'                'Увеличиваем текущее значение максимальной интенсивности
'                If curQ < q0(i) Then
'                    curQprev = curQ
'                    curQ = q0(i)
'                End If
'                If curQ > QQ Then
'                    'Если значение было найдено - возвращаем его
'                    GetVByQ = curQ
'                    Exit Function
'                End If
'            Next i
            GetVByQ = GetVByQSingle(V0, q0, QQ)
        Case Is = wayClasses.door
            GetVByQ = 0
        Case Is = wayClasses.laddo
            GetVByQ = GetVByQSingle(V2, q2, QQ)
        Case Is = wayClasses.ladup
            GetVByQ = GetVByQSingle(V3, q3, QQ)
    End Select
    
'    'Если значение не было найдено - возвращаем максимальное из всего массива
'    GetVByQ = curQ
'
'
'    'Здесь интерполяцию не будем выполнять
'    Select Case wayClass
'        Case Is = wayClasses.hor
'            GetV = Interpolate(DD, Ds(i), Ds(j), V0(i), V0(j))
'        Case Is = wayClasses.door
'            GetV = 0   'Interpolate(DD, Ds(i), Ds(j), V0(i), V0(j))
'        Case Is = wayClasses.laddo
'            GetV = Interpolate(DD, Ds(i), Ds(j), V2(i), V2(j))
'        Case Is = wayClasses.ladup
'            GetV = Interpolate(DD, Ds(i), Ds(j), V3(i), V3(j))
'    End Select
    
    
End Function

Public Function GetVByQSingle(ByRef V_arr As Variant, ByRef Q_arr As Variant, ByVal QQ As Single) As Single
Dim i As Integer
Dim curQprev As Single
Dim curQ As Single
Dim curVprev As Single
Dim curV As Single
Dim isFound As Boolean
    
    For i = 0 To UBound(V_arr)
        'Увеличиваем значение текущего q до тех пор пока следующее в массиве больше его. Если меньше, значит значения пошли на спад - не меняем, сохраняя максимальное
        If curQ < Q_arr(i) Then
            curQprev = curQ
            curQ = Q_arr(i)
            curVprev = curV
            curV = V_arr(i)
        End If
        'Если очередное значение q в массиве оказалось больше искомого - значит мы нашли нужный промежуток
        If curQ > QQ Then
            isFound = True
            Exit For
        End If
    Next i
    
    If isFound Then
        GetVByQSingle = Interpolate(QQ, curQprev, curQ, curVprev, curV)
    Else
        GetVByQSingle = curQ
    End If
End Function
