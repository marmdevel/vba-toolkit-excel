Attribute VB_Name = "modFuncionesPersonalizadas"
Option Explicit

Private Const NOMBRE_PROCEDIMIENTO_PRINCIPAL As String = "AutonumerarColumnaActiva"

' Estructura para almacenar de forma segura el estado previo de Excel
Private Type EstadoEntornoExcel
    screenUpdating As Boolean
    calculation As XlCalculation
    enableEvents As Boolean
    displayAlerts As Boolean
    capturado As Boolean
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

' =============================================================================
' GESTIÓN DE ENTORNO Y UTILIDADES (Compatible 32/64 bits)
' =============================================================================
Private Function CapturarYDesactivarEntorno() As EstadoEntornoExcel
    Dim estadoActual As EstadoEntornoExcel
    
    On Error Resume Next
    estadoActual.screenUpdating = Application.screenUpdating
    estadoActual.calculation = Application.calculation
    estadoActual.enableEvents = Application.enableEvents
    estadoActual.displayAlerts = Application.displayAlerts
    estadoActual.capturado = True
    
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    Application.displayAlerts = False
    On Error GoTo 0
    
    CapturarYDesactivarEntorno = estadoActual
End Function

Private Sub RestaurarEntornoExcel(ByRef estadoAnterior As EstadoEntornoExcel)
    If Not estadoAnterior.capturado Then Exit Sub
    
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

