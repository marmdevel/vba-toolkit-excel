Attribute VB_Name = "modFuncionesPersonalizadas"
Option Explicit

Private Const NOMBRE_PROCEDIMIENTO_PRINCIPAL As String = "AutonumerarColumnaActiva"

' Estructura para almacenar de forma segura el estado previo de Excel
Public Type EstadoEntornoExcel
    screenUpdating As Boolean
    calculation As XlCalculation
    enableEvents As Boolean
    displayAlerts As Boolean
    estadoCapturado As Boolean
End Type

' =============================================================================
' PROCEDIMIENTO PRINCIPAL (Entrada / Interfaz de usuario)
' =============================================================================
Public Sub AutonumerarColumnaActiva()
    Const NOMBRE_PROC As String = NOMBRE_PROCEDIMIENTO_PRINCIPAL
    
    Dim estadoAnterior As EstadoEntornoExcel
    Dim wbActivo As Workbook
    Dim wsActiva As Worksheet
    Dim indiceColumna As Long
    Dim ultimaFila As Long
    Dim rngObjetivo As Range
    Dim datosColumna As Variant
    Dim filasProcesadas As Long
    
    On Error GoTo ManejarError
    
    ' 1. Guardar y optimizar el entorno de Excel antes de modificar nada
    estadoAnterior = CapturarYDesactivarEntorno()
    
    ' 2. Validar precondiciones de libro, hoja, celda activa y protección
    If Not ValidarEntornoColumnaActiva(wbActivo, wsActiva, indiceColumna) Then
        GoTo Salir
    End If
    
    ' 3. Determinar el rango útil de la columna activa (desde fila 1 hasta última fila con datos)
    'ultimaFila = ObtenerUltimaFilaConDatosEnColumna(wsActiva, indiceColumna)
    ultimaFila = get_ObtenerUltimaFilaSpecial(wsActiva)
    
    If ultimaFila < 1 Then
        MsgBox "La columna " & ObtenerLetraColumna(indiceColumna) & _
               " no contiene celdas con datos para evaluar.", _
               vbExclamation, "Autonumeración cancelada"
        GoTo Salir
    End If
    
    Set rngObjetivo = wsActiva.Range(wsActiva.Cells(1, indiceColumna), wsActiva.Cells(ultimaFila, indiceColumna))
    
    ' 4. Leer rango a memoria (si es una sola celda, normalizar a matriz 1x1)
    datosColumna = LeerRangoComoMatriz2D(rngObjetivo)
    
    ' 5. Procesar autonumeración aplicando reglas de negocio en memoria
    filasProcesadas = ProcesarAutonumeracionEnArray(datosColumna)
    
    ' 6. Escribir el resultado de vuelta a la hoja en una sola operación
    rngObjetivo.Value = datosColumna
    
    MsgBox "Autonumeración completada en la columna " & ObtenerLetraColumna(indiceColumna) & _
           " (filas 1 a " & CStr(ultimaFila) & ")." & vbCrLf & _
           "Celdas numeradas/actualizadas: " & CStr(filasProcesadas) & ".", _
           vbInformation, "Proceso finalizado"

Salir:
    Call RestaurarEntornoExcel(estadoAnterior)
    Set rngObjetivo = Nothing
    Set wsActiva = Nothing
    Set wbActivo = Nothing
    Exit Sub

ManejarError:
    Call RestaurarEntornoExcel(estadoAnterior)
    MsgBox "Se produjo un error inesperado durante la ejecución de '" & NOMBRE_PROC & "'." & vbCrLf & vbCrLf & _
           "Número de error: " & CStr(Err.Number) & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en " & NOMBRE_PROC
    Set rngObjetivo = Nothing
    Set wsActiva = Nothing
    Set wbActivo = Nothing
End Sub

' =============================================================================
' VALIDACIONES PREVIAS (Programación defensiva)
' =============================================================================
Private Function ValidarEntornoColumnaActiva( _
    ByRef wbOut As Workbook, _
    ByRef wsOut As Worksheet, _
    ByRef columnaOut As Long) As Boolean
    
    On Error GoTo ManejarError
    
    If ActiveWorkbook Is Nothing Then
        MsgBox "No hay ningún libro de Excel activo.", _
               vbExclamation, "Validación requerida"
        Exit Function
    End If
    Set wbOut = ActiveWorkbook
    
    If wbOut.ActiveSheet Is Nothing Then
        MsgBox "No se detectó una hoja activa en el libro '" & wbOut.Name & "'.", _
               vbExclamation, "Validación requerida"
        Exit Function
    End If
    
    ' Verificar que la hoja activa sea una hoja de cálculo y no un gráfico u otro tipo
    If TypeName(wbOut.ActiveSheet) <> "Worksheet" Then
        MsgBox "La hoja activa debe ser una hoja de cálculo (Worksheet) para autonumerar columnas.", _
               vbExclamation, "Tipo de hoja no compatible"
        Exit Function
    End If
    Set wsOut = wbOut.ActiveSheet
    
    If ActiveCell Is Nothing Then
        MsgBox "Seleccione una celda dentro de la columna que desea autonumerar.", _
               vbExclamation, "Celda activa no disponible"
        Exit Function
    End If
    
    If wsOut.ProtectContents Then
        MsgBox "La hoja '" & wsOut.Name & "' está protegida. Desprotéjala antes de ejecutar la macro.", _
               vbExclamation, "Hoja protegida"
        Exit Function
    End If
    
    columnaOut = ActiveCell.Column
    ValidarEntornoColumnaActiva = True
    Exit Function

ManejarError:
    Err.Raise Err.Number, "ValidarEntornoColumnaActiva", Err.Description
End Function

' =============================================================================
' PROCESAMIENTO DE DATOS Y REGLAS DE NEGOCIO
' =============================================================================
Private Function ObtenerUltimaFilaConDatosEnColumna( _
    ByVal ws As Worksheet, _
    ByVal indiceColumna As Long) As Long
    
    Dim rngColumna As Range
    Dim celdaEncontrada As Range
    
    On Error GoTo ManejarError
    
    Set rngColumna = ws.Columns(indiceColumna)
    
    ' Buscar la última celda con valor visible desde el final hacia arriba
    Set celdaEncontrada = rngColumna.Find( _
        What:="*", _
        After:=rngColumna.Cells(1, 1), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False)
    
    If Not celdaEncontrada Is Nothing Then
        ObtenerUltimaFilaConDatosEnColumna = celdaEncontrada.Row
    Else
        ObtenerUltimaFilaConDatosEnColumna = 0
    End If
    
    Set celdaEncontrada = Nothing
    Set rngColumna = Nothing
    Exit Function

ManejarError:
    Set celdaEncontrada = Nothing
    Set rngColumna = Nothing
    Err.Raise Err.Number, "ObtenerUltimaFilaConDatosEnColumna", Err.Description
End Function

Private Function LeerRangoComoMatriz2D(ByVal rng As Range) As Variant
    Dim matrizTemp As Variant
    Dim matrizNormalizada(1 To 1, 1 To 1) As Variant
    
    On Error GoTo ManejarError
    
    If rng.Cells.CountLarge = 1 Then
        matrizNormalizada(1, 1) = rng.Value
        LeerRangoComoMatriz2D = matrizNormalizada
    Else
        matrizTemp = rng.Value
        LeerRangoComoMatriz2D = matrizTemp
    End If
    Exit Function

ManejarError:
    Err.Raise Err.Number, "LeerRangoComoMatriz2D", Err.Description
End Function

Private Function ProcesarAutonumeracionEnArray(ByRef datosColumna As Variant) As Long
    Dim filaActual As Long
    Dim totalFilas As Long
    Dim contadorNumero As Long
    Dim totalNumeradas As Long
    
    On Error GoTo ManejarError
    
    totalFilas = UBound(datosColumna, 1)
    contadorNumero = 1
    totalNumeradas = 0
    
    For filaActual = 1 To totalFilas
        ' Si la celda tiene un valor numérico (o está vacía), asigna el número correlativo;
        ' si tiene texto, fecha, error o booleano, se deja intacta y continúa.
        If EsValorNumericoReemplazable(datosColumna(filaActual, 1)) Then
            datosColumna(filaActual, 1) = contadorNumero
            contadorNumero = contadorNumero + 1
            totalNumeradas = totalNumeradas + 1
        End If
    Next filaActual
    
    ProcesarAutonumeracionEnArray = totalNumeradas
    Exit Function

ManejarError:
    Err.Raise Err.Number, "ProcesarAutonumeracionEnArray", Err.Description
End Function

Private Function EsValorNumericoReemplazable(ByVal valorCelda As Variant) As Boolean
    Dim tipoVar As VbVarType
    
    ' Los valores de error de Excel (ej. #N/A, #VALUE!) no son numéricos
    If IsError(valorCelda) Then
        EsValorNumericoReemplazable = False
        Exit Function
    End If
    
    ' Los valores booleanos (VERDADERO/FALSO) no deben tratarse como números de secuencia
    If VarType(valorCelda) = vbBoolean Then
        EsValorNumericoReemplazable = False
        Exit Function
    End If
    
    ' Celdas vacías o valores numéricos reales/numéricos válidos
    If IsEmpty(valorCelda) Then
        EsValorNumericoReemplazable = True
        Exit Function
    End If
    
    tipoVar = VarType(valorCelda)
    If tipoVar = vbInteger Or tipoVar = vbLong Or tipoVar = vbSingle Or _
       tipoVar = vbDouble Or tipoVar = vbCurrency Or tipoVar = vbDecimal Then
        EsValorNumericoReemplazable = True
        Exit Function
    End If
    
    ' Si es texto, verificar si representa un número válido (ej. "10" o "3.5")
    If tipoVar = vbString Then
        If Trim(valorCelda) <> "" And IsNumeric(valorCelda) Then
            EsValorNumericoReemplazable = True
            Exit Function
        End If
    End If
    
    EsValorNumericoReemplazable = False
End Function



Public Sub RestaurarEntornoExcel(ByRef estadoAnterior As EstadoEntornoExcel)
    If Not estadoAnterior.estadoCapturado Then Exit Sub
    
    On Error Resume Next
    Application.screenUpdating = estadoAnterior.screenUpdating
    Application.calculation = estadoAnterior.calculation
    Application.enableEvents = estadoAnterior.enableEvents
    Application.displayAlerts = estadoAnterior.displayAlerts
    On Error GoTo 0
End Sub

Private Function ObtenerLetraColumna(ByVal indiceColumna As Long) As String
    Dim partesNombre() As String
    partesNombre = Split(Cells(1, indiceColumna).Address(True, False), "$")
    ObtenerLetraColumna = partesNombre(0)
End Function



' =============================================================================
' rutina: modOrdenarHojaPorColumna
' Propósito: Ordenar la hoja activa completa de forma ascendente según la
'            columna de la celda activa o las dos primeras columnas únicas
'            seleccionadas (ordenadas de izquierda a derecha, contiguas o no).
' Entorno:   Excel 2019 (64 bits / 32 bits compatible).
' =============================================================================

' -----------------------------------------------------------------------------
' Procedimiento principal (interfaz de ejecución)
' -----------------------------------------------------------------------------
Public Sub OrdenarHojaPorColumnaActiva(Optional ByVal ordenAscendente As String = "asc")
    Const NOMBRE_PROC As String = "OrdenarHojaPorColumnaActiva"

    Dim hojaActiva As Worksheet
    Dim estadoPrevio As EstadoEntornoExcel
    Dim columnasClave() As Long
    Dim cantidadClaves As Long
    Dim mensajeValidacion As String
    Dim ordenXl As XlSortOrder

    On Error GoTo ManejarError

    ' 1. Validar precondiciones antes de modificar el estado de Excel
    If Not ValidarPrecondicionesOrden(hojaActiva, columnasClave, cantidadClaves, mensajeValidacion) Then
        MsgBox mensajeValidacion, vbExclamation, "Validación de ordenamiento"
        GoTo Salir
    End If

    ' 2. Capturar y optimizar propiedades del entorno de Excel
    estadoPrevio = CapturarYDesactivarEntorno()

    If ordenAscendente = "asc" Then
        ordenXl = xlAscending
    Else
        ordenXl = xlDescending
    End If

    ' 3. Ejecutar el ordenamiento sobre la hoja completa
    AplicarOrdenEnHojaCompleta hojaActiva, columnasClave, cantidadClaves, ordenXl

    MsgBox "La hoja '" & hojaActiva.Name & "' se ordenó correctamente.", vbInformation, "Ordenamiento completado"

Salir:
    RestaurarEntorno estadoPrevio
    Set hojaActiva = Nothing
    Exit Sub

ManejarError:
    RestaurarEntorno estadoPrevio
    MsgBox "Ocurrió un error al ordenar la hoja." & vbCrLf & vbCrLf & _
           "Procedimiento: " & NOMBRE_PROC & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en ejecución"
    Set hojaActiva = Nothing
End Sub

' -----------------------------------------------------------------------------
' Validación defensiva de hoja, selección y columnas clave
' -----------------------------------------------------------------------------
Private Function ValidarPrecondicionesOrden( _
    ByRef hojaDestino As Worksheet, _
    ByRef columnasClave() As Long, _
    ByRef cantidadClaves As Long, _
    ByRef mensajeError As String _
) As Boolean
    Dim rangoSeleccionado As Range
    Dim ultimaFila As Long
    Dim ultimaColumna As Long

    ValidarPrecondicionesOrden = False
    cantidadClaves = 0

    If ActiveWorkbook Is Nothing Then
        mensajeError = "No hay un libro de Excel activo."
        Exit Function
    End If

    If Not TypeOf ActiveSheet Is Worksheet Then
        mensajeError = "La hoja activa debe ser una hoja de cálculo válida."
        Exit Function
    End If

    Set hojaDestino = ActiveSheet

    If hojaDestino.ProtectContents Then
        mensajeError = "La hoja '" & hojaDestino.Name & "' está protegida. Desprotégela antes de ordenar."
        Exit Function
    End If

    If ActiveCell Is Nothing Or Not (TypeOf Selection Is Range) Then
        mensajeError = "Selecciona una celda o un rango válido dentro de la hoja antes de ejecutar el orden."
        Exit Function
    End If

    Set rangoSeleccionado = Selection

    If Not ObtenerUltimaCeldaUsada(hojaDestino, ultimaFila, ultimaColumna) Then
        mensajeError = "La hoja '" & hojaDestino.Name & "' no contiene datos para ordenar."
        Exit Function
    End If

    ' Extraer hasta 2 columnas únicas ordenadas de izquierda a derecha
    ObtenerColumnasOrdenamiento rangoSeleccionado, ActiveCell.Column, ultimaColumna, columnasClave, cantidadClaves

    If cantidadClaves <= 0 Then
        mensajeError = "No se pudo determinar una columna válida para ordenar."
        Exit Function
    End If

    ValidarPrecondicionesOrden = True
End Function



' -----------------------------------------------------------------------------
' Obtiene las columnas únicas de la selección (máx. 2), ordenadas por índice
' Si solo hay 1 celda activa, usa su columna; si hay >1 columna, toma las
' dos primeras según orden de columna (ej. B y A -> A luego B).
' -----------------------------------------------------------------------------
Private Sub ObtenerColumnasOrdenamiento( _
    ByVal seleccionActual As Range, _
    ByVal columnaCeldaActiva As Long, _
    ByVal ultimaColumnaHoja As Long, _
    ByRef columnasOrdenadas() As Long, _
    ByRef totalColumnas As Long _
)
    Dim areaActual As Range
    Dim colIdx As Long
    Dim tempCols() As Long
    Dim countTemp As Long
    Dim i As Long
    Dim j As Long
    Dim swapVal As Long

    totalColumnas = 0
    countTemp = 0

    ' Si la selección es una única celda, prioriza la columna de la celda activa
    If seleccionActual.CountLarge = 1 Then
        If columnaCeldaActiva >= 1 And columnaCeldaActiva <= ultimaColumnaHoja Then
            ReDim columnasOrdenadas(1 To 1)
            columnasOrdenadas(1) = columnaCeldaActiva
            totalColumnas = 1
        End If
        Exit Sub
    End If

    ' Recolectar columnas únicas presentes en todas las áreas de la selección
    ReDim tempCols(1 To seleccionActual.Columns.Count + 1)

    For Each areaActual In seleccionActual.Areas
        For colIdx = areaActual.Column To areaActual.Column + areaActual.Columns.Count - 1
            If colIdx >= 1 And colIdx <= ultimaColumnaHoja Then
                If Not ExisteColumnaEnArray(tempCols, countTemp, colIdx) Then
                    countTemp = countTemp + 1
                    If countTemp > UBound(tempCols) Then
                        ReDim Preserve tempCols(1 To countTemp)
                    End If
                    tempCols(countTemp) = colIdx
                End If
            End If
        Next colIdx
    Next areaActual

    ' Respaldo si ninguna columna de la selección cae dentro del rango usado
    If countTemp = 0 And columnaCeldaActiva >= 1 And columnaCeldaActiva <= ultimaColumnaHoja Then
        countTemp = 1
        tempCols(1) = columnaCeldaActiva
    End If

    ' Ordenar índices de columna de menor a mayor (ej. Columna A = 1 antes que Columna B = 2)
    For i = 1 To countTemp - 1
        For j = i + 1 To countTemp
            If tempCols(j) < tempCols(i) Then
                swapVal = tempCols(i)
                tempCols(i) = tempCols(j)
                tempCols(j) = swapVal
            End If
        Next j
    Next i

    ' Limitar a máximo 2 columnas clave según regla de negocio
    If countTemp >= 2 Then
        totalColumnas = 2
        ReDim columnasOrdenadas(1 To 2)
        columnasOrdenadas(1) = tempCols(1)
        columnasOrdenadas(2) = tempCols(2)
    ElseIf countTemp = 1 Then
        totalColumnas = 1
        ReDim columnasOrdenadas(1 To 1)
        columnasOrdenadas(1) = tempCols(1)
    End If
End Sub

Private Function ExisteColumnaEnArray( _
    ByRef arrCols() As Long, _
    ByVal limite As Long, _
    ByVal valorBuscado As Long _
) As Boolean
    Dim idx As Long

    For idx = 1 To limite
        If arrCols(idx) = valorBuscado Then
            ExisteColumnaEnArray = True
            Exit Function
        End If
    Next idx

    ExisteColumnaEnArray = False
End Function

' -----------------------------------------------------------------------------
' Aplica el ordenamiento sobre la hoja completa usando objeto Sort de Excel
' -----------------------------------------------------------------------------
Private Sub AplicarOrdenEnHojaCompleta( _
    ByVal hoja As Worksheet, _
    ByRef columnasClave() As Long, _
    ByVal cantidadClaves As Long, _
    ByVal direccionOrden As XlSortOrder _
)
    Const FILA_INICIO_DATOS As Long = 2
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoDatos As Range
    Dim campoClave1 As Range
    Dim campoClave2 As Range

    If Not ObtenerUltimaCeldaUsada(hoja, ultimaFila, ultimaColumna) Then
        Exit Sub
    End If

    ' Si solo hay encabezado (fila 1) o no hay filas de datos, no hay nada que ordenar
    If ultimaFila < FILA_INICIO_DATOS Then
        Exit Sub
    End If

    Set rangoDatos = hoja.Range( _
        hoja.Cells(FILA_INICIO_DATOS, 1), _
        hoja.Cells(ultimaFila, ultimaColumna) _
    )
    Set campoClave1 = hoja.Range( _
        hoja.Cells(FILA_INICIO_DATOS, columnasClave(1)), _
        hoja.Cells(ultimaFila, columnasClave(1)) _
    )

    With hoja.Sort
        .SortFields.Clear
        .SortFields.Add2 _
            key:=campoClave1, _
            SortOn:=xlSortOnValues, _
            Order:=direccionOrden, _
            DataOption:=xlSortNormal

        If cantidadClaves >= 2 Then
            Set campoClave2 = hoja.Range( _
                hoja.Cells(FILA_INICIO_DATOS, columnasClave(2)), _
                hoja.Cells(ultimaFila, columnasClave(2)) _
            )
            .SortFields.Add2 _
                key:=campoClave2, _
                SortOn:=xlSortOnValues, _
                Order:=direccionOrden, _
                DataOption:=xlSortNormal
        End If

        .SetRange rangoDatos
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Public Sub RestaurarEntorno(ByRef estado As EstadoEntornoExcel)
    If Not estado.estadoCapturado Then
        Exit Sub
    End If

    On Error Resume Next
    Application.screenUpdating = estado.screenUpdating
    Application.calculation = estado.calculation
    Application.enableEvents = estado.enableEvents
    Application.displayAlerts = estado.displayAlerts
    On Error GoTo 0
End Sub
' -----------------------------------------------------------------------------
' Determina la última fila y última columna con contenido real en la hoja
' -----------------------------------------------------------------------------
Public Function ObtenerUltimaCeldaUsada( _
    ByVal hoja As Worksheet, _
    ByRef ultimaFila As Long, _
    ByRef ultimaColumna As Long _
) As Boolean
    Dim celdaFinal As Range

    On Error Resume Next
    Set celdaFinal = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )
    On Error GoTo 0

    If celdaFinal Is Nothing Then
        ultimaFila = 0
        ultimaColumna = 0
        ObtenerUltimaCeldaUsada = False
        Exit Function
    End If

    ultimaFila = celdaFinal.Row

    On Error Resume Next
    Set celdaFinal = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )
    On Error GoTo 0

    If celdaFinal Is Nothing Then
        ultimaColumna = 0
        ObtenerUltimaCeldaUsada = False
        Exit Function
    End If

    ultimaColumna = celdaFinal.Column
    ObtenerUltimaCeldaUsada = (ultimaFila >= 1 And ultimaColumna >= 1)
End Function

' -----------------------------------------------------------------------------
' Gestión del estado del entorno de Excel
' -----------------------------------------------------------------------------
Public Function CapturarYDesactivarEntorno() As EstadoEntornoExcel
    Dim estado As EstadoEntornoExcel

    estado.screenUpdating = Application.screenUpdating
    estado.calculation = Application.calculation
    estado.enableEvents = Application.enableEvents
    estado.displayAlerts = Application.displayAlerts
    estado.estadoCapturado = True

    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    Application.displayAlerts = False

    CapturarYDesactivarEntorno = estado
End Function

