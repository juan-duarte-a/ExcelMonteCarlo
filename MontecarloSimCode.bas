Attribute VB_Name = "MontecarloSimCode"
Option Explicit

'Deficiones generales
    Const Columna As Integer = 0
'Fin definiciones generales

Function generar_aleat(ByVal semilla As Double, ByVal min As Double, ByVal max As Double, ByVal entero As Boolean, ByVal r As Boolean) As Double
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    If entero = True Then
        generar_aleat = Int(Rnd() * (max - min + 1) + min)
    Else
        generar_aleat = Rnd() * (max - min) + min
    End If
End Function

Sub secuencia_aleat_uniforme(ByRef sec() As Double, ByVal cantidad As Long, ByVal semilla As Double, ByVal min As Double, ByVal max As Double, _
        ByVal repetir As Boolean, ByVal entero As Boolean, ByVal r As Boolean)
    Dim i As Long
    Dim rnd_number As Double
    
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    If repetir = True Then
        Rnd (-1)
    End If
    
    For i = 1 To cantidad
        rnd_number = generar_aleat(semilla, min, max, entero, r)
        sec(i) = rnd_number
    Next i
End Sub

Sub generar_secuencia_uniforme(ByVal r As Boolean)
Attribute generar_secuencia_uniforme.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim cantidad As Long, semilla As Double, min As Double, max As Double, repetir As Boolean, entero As Boolean
    Dim secuencia() As Double
    Dim i As Long
    
    cantidad = Worksheets("Simulación individual").Cells(6, 3).Value
    ReDim secuencia(1 To cantidad)
    
    semilla = Worksheets("Simulación individual").Cells(2, 3).Value
    min = Worksheets("Simulación individual").Cells(3, 3).Value
    max = Worksheets("Simulación individual").Cells(4, 3).Value
    repetir = Worksheets("Simulación individual").Cells(7, 3)
    entero = Worksheets("Simulación individual").Cells(8, 3)
    
    Call secuencia_aleat_uniforme(secuencia, cantidad, semilla, min, max, repetir, entero, r)

    For i = 1 To cantidad
        Worksheets("Simulación individual").Cells(12 + i, 2) = i
        Worksheets("Simulación individual").Cells(12 + i, 3 + Columna) = secuencia(i)
    Next i
End Sub

Function generar_aleat_triangular(ByVal semilla As Double, ByVal min As Double, ByVal max As Double, ByVal probable As Double, ByVal entero As Boolean, ByVal r As Boolean) As Double
    Dim rnd_number As Double, f_c As Double
    
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    rnd_number = Rnd()
    
    If entero = True Then
        max = max + 1
        f_c = (probable - min) / (max - min)
        If rnd_number > 0 And rnd_number < f_c Then
            generar_aleat_triangular = Int(min + Sqr(rnd_number * (max - min) * (probable - min)))
        Else
            generar_aleat_triangular = Int(max - Sqr((1 - rnd_number) * (max - min) * (max - probable)))
        End If
    Else
        f_c = (probable - min) / (max - min)
        If rnd_number > 0 And rnd_number < f_c Then
            generar_aleat_triangular = min + Sqr(rnd_number * (max - min) * (probable - min))
        Else
            generar_aleat_triangular = max - Sqr((1 - rnd_number) * (max - min) * (max - probable))
        End If
    End If
End Function

Sub secuencia_aleat_triangular(ByRef sec() As Double, ByVal cantidad As Long, ByVal semilla As Double, ByVal min As Double, ByVal max As Double, _
        ByVal probable As Double, ByVal repetir As Boolean, ByVal entero As Boolean, ByVal r As Boolean)
    Dim i As Long
    Dim rnd_number As Double
    
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    If repetir = True Then
        Rnd (-1)
    End If
    
    For i = 1 To cantidad
        rnd_number = generar_aleat_triangular(semilla, min, max, probable, entero, r)
        sec(i) = rnd_number
    Next i
End Sub

Sub generar_secuencia_triangular(ByVal r As Boolean)
    Dim cantidad As Long, semilla As Double, min As Double, max As Double, probable As Double, repetir As Boolean, entero As Boolean
    Dim secuencia() As Double
    Dim i As Long
    
    cantidad = Worksheets("Simulación individual").Cells(6, 3).Value
    ReDim secuencia(1 To cantidad)
    
    semilla = Worksheets("Simulación individual").Cells(2, 3).Value
    min = Worksheets("Simulación individual").Cells(3, 3).Value
    max = Worksheets("Simulación individual").Cells(4, 3).Value
    probable = Worksheets("Simulación individual").Cells(5, 3).Value
    repetir = Worksheets("Simulación individual").Cells(7, 3)
    entero = Worksheets("Simulación individual").Cells(8, 3)
    
    Call secuencia_aleat_triangular(secuencia, cantidad, semilla, min, max, probable, repetir, entero, r)
    
    For i = 1 To cantidad
        Worksheets("Simulación individual").Cells(12 + i, 2) = i
        Worksheets("Simulación individual").Cells(12 + i, 3 + Columna) = secuencia(i)
    Next i
End Sub

Public Sub generar_secuencia()
Attribute generar_secuencia.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim dist As String * 1
    Dim semilla As Double
    Dim aleat As Boolean
    
    dist = Worksheets("Simulación individual").Cells(9, 3).Value
    semilla = Worksheets("Simulación individual").Cells(2, 3).Value
    aleat = Worksheets("Simulación individual").Cells(10, 3).Value
    
    Application.ScreenUpdating = False
    Select Case dist
        Case "U"
            Call generar_secuencia_uniforme(aleat)
        Case "T"
            Call generar_secuencia_triangular(aleat)
        Case "N"
            Call generar_secuencia_normal(aleat)
        End Select
    Application.ScreenUpdating = True
End Sub

Function generar_aleat_normal(ByVal semilla As Double, ByVal media As Double, ByVal desv As Double, ByVal entero As Boolean, ByVal r As Boolean) As Double
    Dim U As Double, V As Double, X As Double, X2 As Double
    Dim calculando As Boolean
    
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    calculando = True
    Do While calculando
        U = Rnd()
        V = Rnd()
        
        X = Sqr(8 / Exp(1)) * (V - 0.5) / U
        X2 = X * X
        
        If X2 <= 5 - 4 * Exp(1 / 4) * U Then
            calculando = False
            If entero = True Then
                generar_aleat_normal = Int(X * desv + media)
            Else
                generar_aleat_normal = X * desv + media
            End If
        ElseIf Not (X2 >= 4 * Exp(-1.35) / U + 1.4) Then
            If X2 <= -4 * Log(U) Then
                calculando = False
                If entero = True Then
                    generar_aleat_normal = Int(X * desv + media)
                Else
                    generar_aleat_normal = X * desv + media
                End If
            End If
        End If
    Loop
End Function

Sub secuencia_aleat_normal(ByRef sec() As Double, ByVal cantidad As Long, ByVal semilla As Double, ByVal media As Double, ByVal desv As Double, _
        ByVal repetir As Boolean, ByVal entero As Boolean, ByVal r As Boolean)
    Dim i As Long
    Dim rnd_number As Double
    
    If r Then
        Randomize
    Else
        Randomize (semilla)
    End If
    
    If repetir = True Then
        Rnd (-1)
    End If
    
    For i = 1 To cantidad
        rnd_number = generar_aleat_normal(semilla, media, desv, entero, r)
        sec(i) = rnd_number
    Next i
End Sub

Sub generar_secuencia_normal(ByVal r As Boolean)
    Dim cantidad As Long, semilla As Double, media As Double, desv As Double, repetir As Boolean, entero As Boolean
    Dim secuencia() As Double
    Dim i As Long
    
    cantidad = Worksheets("Simulación individual").Cells(6, 3).Value
    ReDim secuencia(1 To cantidad)
    
    semilla = Worksheets("Simulación individual").Cells(2, 3).Value
    media = Worksheets("Simulación individual").Cells(3, 3).Value
    desv = Worksheets("Simulación individual").Cells(4, 3).Value
    repetir = Worksheets("Simulación individual").Cells(7, 3)
    entero = Worksheets("Simulación individual").Cells(8, 3)
    
    Call secuencia_aleat_normal(secuencia, cantidad, semilla, media, desv, repetir, entero, r)
    
    For i = 1 To cantidad
        Worksheets("Simulación individual").Cells(12 + i, 2) = i
        Worksheets("Simulación individual").Cells(12 + i, 3 + Columna) = secuencia(i)
    Next i
End Sub

Sub sim_montecarlo()
Attribute sim_montecarlo.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim i As Long, j As Long, k As Integer, l As Integer, n As Integer, o As Integer, mostrar_datos As Boolean
    Dim iteraciones As Long, iter_externa As Long, cant_ext As Integer
    Dim min() As Double, max() As Double, promedio() As Double, desv() As Double
    Dim secuencia() As Double, val_obs_externa() As Double
    Dim tipo_var As String * 1, dist_prob_var As String * 1
    Dim variables() As Variant, variables_ext() As Integer
    Dim observaciones() As Variant, val_observaciones() As Double
    Dim aleatorios_int() As Double, aleatorios_ext() As Double
    
    Application.StatusBar = "Iniciando simulación..."
    mostrar_datos = Worksheets("Parámetros").Cells(5, 4).Value
    iteraciones = Worksheets("Parámetros").Cells(4, 4).Value
    iter_externa = Worksheets("Parámetros").Cells(4, 6).Value

    i = 1
    n = 0
    cant_ext = 0
    Do While i <= 40
        If Worksheets("Parámetros").Cells(8 + i, 4).Value = "S" Then
            ReDim Preserve variables(1 To 12, 1 To n + 1)
            For j = 1 To 11
                variables(j, n + 1) = Worksheets("Parámetros").Cells(8 + i, 2 + j).Value
                If j = 3 Then
                    If variables(j, n + 1) = "E" Then
                        cant_ext = cant_ext + 1
                        ReDim Preserve variables_ext(1 To cant_ext)
                        variables_ext(cant_ext) = n + 1
                    End If
                End If
            Next j
            variables(12, n + 1) = 8 + i
            n = n + 1
        End If
        i = i + 1
    Loop

    i = 1
    o = 0
    Do While i <= 40
        If Worksheets("Parámetros").Cells(52 + i, 4).Value = "S" Then
            ReDim Preserve observaciones(1 To 7, 1 To o + 1)
            For j = 1 To 2
                observaciones(j, o + 1) = Worksheets("Parámetros").Cells(52 + i, 2 + j).Value
            Next j
            observaciones(7, o + 1) = i
            o = o + 1
        End If
        i = i + 1
    Loop

    Application.StatusBar = "Generando aleatorios..."
    ReDim aleatorios_int(1 To n - cant_ext, 1 To iteraciones)
    If cant_ext > 0 Then ReDim aleatorios_ext(1 To cant_ext, 1 To iter_externa)
    ReDim secuencia(1 To WorksheetFunction.max(iteraciones, iter_externa))
    ReDim val_observaciones(1 To o, 1 To WorksheetFunction.max(iteraciones, iter_externa))
    ReDim val_obs_externa(1 To o, 1 To iter_externa)
    ReDim min(1 To o)
    ReDim max(1 To o)
    ReDim promedio(1 To o)
    ReDim desv(1 To o)
    
    Application.ScreenUpdating = False
    Worksheets("Parámetros").Range("N9:N48").ClearContents
    Worksheets("Datos").Cells.ClearContents

    If cant_ext > 0 Then
        For i = 1 To cant_ext
            tipo_var = variables(4, variables_ext(i))
            Select Case tipo_var
                Case "U"
                    Call secuencia_aleat_uniforme(secuencia, iter_externa, _
                            variables(5, variables_ext(i)), variables(6, variables_ext(i)), variables(7, variables_ext(i)), variables(9, variables_ext(i)), variables(10, variables_ext(i)), variables(11, variables_ext(i)))
                Case "T"
                    Call secuencia_aleat_triangular(secuencia, iter_externa, _
                            variables(5, variables_ext(i)), variables(6, variables_ext(i)), variables(7, variables_ext(i)), variables(8, variables_ext(i)), variables(9, variables_ext(i)), variables(10, variables_ext(i)), variables(11, variables_ext(i)))
                Case "N"
                    Call secuencia_aleat_normal(secuencia, iter_externa, _
                            variables(5, variables_ext(i)), variables(6, variables_ext(i)), variables(7, variables_ext(i)), variables(9, variables_ext(i)), variables(10, variables_ext(i)), variables(11, variables_ext(i)))
            End Select
            
            For j = 1 To iter_externa
                aleatorios_ext(i, j) = secuencia(j)
            Next j
        Next i
        
        For i = 1 To iter_externa
            j = 0
            For k = 1 To n
                If variables(3, k) = "I" Then
                    j = j + 1
                    tipo_var = variables(4, k)
                    Select Case tipo_var
                        Case "U"
                            Call secuencia_aleat_uniforme(secuencia, iteraciones, _
                                    variables(5, k), variables(6, k), variables(7, k), variables(9, k), variables(10, k), variables(11, k))
                        Case "T"
                            Call secuencia_aleat_triangular(secuencia, iteraciones, _
                                    variables(5, k), variables(6, k), variables(7, k), variables(8, k), variables(9, k), variables(10, k), variables(11, k))
                        Case "N"
                            Call secuencia_aleat_normal(secuencia, iteraciones, _
                                    variables(5, k), variables(6, k), variables(7, k), variables(9, k), variables(10, k), variables(11, k))
                    End Select
                    
                    For l = 1 To iteraciones
                        aleatorios_int(j, l) = secuencia(l)
                    Next l
                End If
            Next k
            
            Call actualizar_variables(iteraciones, n, o, variables, aleatorios_int, aleatorios_ext, observaciones, val_observaciones, cant_ext, i, iter_externa)
            Call calc_estadísticas(o, iteraciones, val_observaciones, observaciones)
        
            For l = 1 To o
                val_obs_externa(l, i) = observaciones(5, l)
                If i = 1 Then
                    min(l) = observaciones(5, l)
                    max(l) = observaciones(5, l)
                    promedio(l) = observaciones(5, l)
                    desv(l) = observaciones(5, l) ^ 2
                Else
                    If observaciones(5, l) < min(l) Then min(l) = observaciones(5, l)
                    If observaciones(5, l) > max(l) Then max(l) = observaciones(5, l)
                    promedio(l) = (promedio(l) * (i - 1) + observaciones(5, l)) / i
                    desv(l) = desv(l) + observaciones(5, l) ^ 2
                End If
            Next l
            
            If mostrar_datos = True Then
                For k = 1 To cant_ext
                    Worksheets("Datos").Cells(2 + i, 2 + k) = aleatorios_ext(k, i)
                Next k
            End If
        Next i
        
        For i = 1 To o
            desv(i) = Sqr((desv(i) / iter_externa - promedio(i) ^ 2) * iter_externa / (iter_externa - 1))
            observaciones(3, i) = min(i)
            observaciones(4, i) = max(i)
            observaciones(5, i) = promedio(i)
            observaciones(6, i) = desv(i)
        Next i
    Else
        For i = 1 To n
            tipo_var = variables(4, i)
            Select Case tipo_var
                Case "U"
                    Call secuencia_aleat_uniforme(secuencia, iteraciones, _
                            variables(5, i), variables(6, i), variables(7, i), variables(9, i), variables(10, i), variables(11, i))
                Case "T"
                    Call secuencia_aleat_triangular(secuencia, iteraciones, _
                            variables(5, i), variables(6, i), variables(7, i), variables(8, i), variables(9, i), variables(10, i), variables(11, i))
                Case "N"
                    Call secuencia_aleat_normal(secuencia, iteraciones, _
                            variables(5, i), variables(6, i), variables(7, i), variables(9, i), variables(10, i), variables(11, i))
            End Select
            
            For j = 1 To iteraciones
                aleatorios_int(i, j) = secuencia(j)
            Next j
        Next i
        
        Call actualizar_variables(iteraciones, n, o, variables, aleatorios_int, aleatorios_ext, observaciones, val_observaciones, cant_ext, 0, 0)
        Application.StatusBar = "Calculando estadísticas..."
        Call calc_estadísticas(o, iteraciones, val_observaciones, observaciones)
    End If
    
    Worksheets("Resultados").Range("C5:G34").ClearContents
    
    For i = 1 To o
        Worksheets("Resultados").Cells(4 + i, 3) = observaciones(1, i)
        For j = 3 To 6
            Worksheets("Resultados").Cells(4 + i, 2 + j - 1) = observaciones(j, i)
        Next j
    Next i
    
    If mostrar_datos = True Then
        Application.StatusBar = "Llenando hoja con valores de variables observadas..."
        Application.Calculation = xlManual
        Application.ScreenUpdating = False
    
        If cant_ext = 0 Then
            Worksheets("Datos").Cells(2, 2).FormulaR1C1 = "Iteración"
            Worksheets("Datos").Cells(2, 3).FormulaR1C1 = "Observación 1"
            Worksheets("Datos").Range("C2").AutoFill _
                Destination:=Worksheets("Datos").Range(Worksheets("Datos").Cells(2, 3).Address, Worksheets("Datos").Cells(2, o + 2).Address), Type:=xlFillDefault
            
            For j = 1 To iteraciones
                Worksheets("Datos").Cells(2 + j, 2) = j
            Next j
            
            For i = 1 To o
                For j = 1 To iteraciones
                    Worksheets("Datos").Cells(2 + j, 2 + i) = val_observaciones(i, j)
                Next j
            Next i
        Else
            Worksheets("Datos").Cells(2, 2).FormulaR1C1 = "Iteración"
            For i = 1 To cant_ext
                Worksheets("Datos").Cells(2, 2 + i) = "Variable ext. " & variables_ext(i)
            Next i
            
            Worksheets("Datos").Cells(2, 3 + cant_ext).FormulaR1C1 = "Observación 1"
            Worksheets("Datos").Range(Cells(2, 3 + cant_ext).Address).AutoFill _
                Destination:=Worksheets("Datos").Range(Worksheets("Datos").Cells(2, 3 + cant_ext).Address, Worksheets("Datos").Cells(2, 2 + cant_ext + o).Address), Type:=xlFillDefault
            
            For j = 1 To iter_externa
                Worksheets("Datos").Cells(2 + j, 2) = j
            Next j
            
            For i = 1 To o
                For j = 1 To iter_externa
                    Worksheets("Datos").Cells(2 + j, 2 + i + cant_ext) = val_obs_externa(i, j)
                Next j
            Next i
        End If
        
        Application.Calculation = xlAutomatic
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Simulación finalizada."
End Sub

Sub actualizar_variables(ByVal iteraciones As Integer, ByVal n As Integer, ByVal o As Integer, variables() As Variant, aleatorios_int() As Double, _
        ByRef aleatorios_ext() As Double, ByRef observaciones() As Variant, ByRef val_observaciones() As Double, cant_ext As Integer, i_ext As Long, iter_externa As Long)
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim t_inicio As Double, t_fin As Double
    Dim str_statusbar As String
    
    t_fin = 0.5
    t_inicio = Timer
    
    Application.Calculation = xlManual
    
    For i = 1 To iteraciones
        k = 0
        l = 0
        For j = 1 To n
            If variables(3, j) = "I" Then
                k = k + 1
                Worksheets("Parámetros").Cells(variables(12, j), 14) = aleatorios_int(k, i)
            Else
                l = l + 1
                Worksheets("Parámetros").Cells(variables(12, j), 14) = aleatorios_ext(l, i_ext)
            End If
        Next j
        
        Calculate

        For k = 1 To o
            val_observaciones(k, i) = Worksheets("Parámetros").Cells(52 + observaciones(7, k), 5).Value
        Next k
        
        If Timer >= t_inicio + t_fin Then
            Application.ScreenUpdating = True
            If cant_ext = 0 Then
                str_statusbar = "Actualizando hoja de cálculo: Iteración " & i & " / " & iteraciones
            Else
                str_statusbar = "Actualizando hoja de cálculo: Iteración externa " _
                    & i_ext & " / " & iter_externa & ", Iteración interna " & i & " / " & iteraciones
            End If
            Application.StatusBar = str_statusbar
            t_inicio = Timer
        Else
            If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
        End If
    Next i
    
    Application.Calculation = xlAutomatic
End Sub

Sub calc_estadísticas(ByVal o As Integer, ByVal iteraciones As Long, ByRef val_observaciones() As Double, ByRef observaciones() As Variant)
    Dim i As Integer, j As Integer
    Dim min As Double, max As Double, promedio As Double, desv As Double
    
    For i = 1 To o
        min = val_observaciones(i, 1)
        max = val_observaciones(i, 1)
        promedio = 0
        desv = 0
        
        For j = 1 To iteraciones
            If val_observaciones(i, j) < min Then min = val_observaciones(i, j)
            If val_observaciones(i, j) > max Then max = val_observaciones(i, j)
            promedio = (promedio * (j - 1) + val_observaciones(i, j)) / j
            desv = desv + val_observaciones(i, j) ^ 2
        Next j
        
        On Error Resume Next
        desv = Sqr((desv / iteraciones - promedio ^ 2) * iteraciones / (iteraciones - 1))
        
        observaciones(3, i) = min
        observaciones(4, i) = max
        observaciones(5, i) = promedio
        observaciones(6, i) = desv
    Next i
End Sub
