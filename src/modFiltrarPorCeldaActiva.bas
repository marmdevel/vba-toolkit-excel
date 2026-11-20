Attribute VB_Name = "modFiltrarPorCeldaActiva"


' =============================================================================
' Módulo: modFiltrarPorCeldaActiva
' Propósito: Filtra mediante AutoFilter la tabla (ListObject) o el rango de
'            datos de la hoja activa según el valor y columna de la celda activa,
'            considerando posibles columnas iniciales en blanco.
' Entorno:   Excel 2019 (64 bits / 32 bits compatible).
' =============================================================================

' -----------------------------------------------------------------------------
' Procedimiento principal de entrada
' -----------------------------------------------------------------------------
Public Sub FiltrarPorCeldaActiva()
    Const NOMBRE_PROC As String = "FiltrarPorCeldaActiva"

    Dim hojaActiva As Worksheet
    Dim celdaActiva As Range
    Dim tablaActiva As ListObject
    Dim rangoFiltro As Range
    Dim campoRelativo As Long
    Dim estadoPrevio As EstadoEntornoExcel
    Dim mensajeValidacion As String

    On Error GoTo ManejarError

    ' 1. Validar precondiciones antes de alterar el estado de Excel
    If Not ValidarPrecondicionesFiltro(hojaActiva, celdaActiva, tablaActiva, mensajeValidacion) Then
        MsgBox mensajeValidacion, vbExclamation, "Validación de filtro"
        GoTo Salir
    End If

    ' 2. Capturar y optimizar propiedades del entorno de Excel
    estadoPrevio = CapturarYDesactivarEntorno()

    ' 3. Resolver el rango objetivo y el número de campo relativo
    If Not tablaActiva Is Nothing Then
        Set rangoFiltro = tablaActiva.Range
        campoRelativo = celdaActiva.Column - rangoFiltro.Column + 1
    Else
        If Not ResolverRangoHojaConColumnasVacias(hojaActiva, celdaActiva, rangoFiltro, campoRelativo, mensajeValidacion) Then
            MsgBox mensajeValidacion, vbExclamation, "Validación de rango"
            GoTo Salir
        End If
    End If

    ' 4. Aplicar el autofiltro con el valor de la celda activa
    AplicarAutoFiltroEnRango hojaActiva, rangoFiltro, campoRelativo, celdaActiva

Salir:
    RestaurarEntorno estadoPrevio
    Set rangoFiltro = Nothing
    Set tablaActiva = Nothing
    Set celdaActiva = Nothing
    Set hojaActiva = Nothing
    Exit Sub

ManejarError:
    RestaurarEntorno estadoPrevio
    MsgBox "No fue posible aplicar el filtro por la celda activa." & vbCrLf & vbCrLf & _
           "Procedimiento: " & NOMBRE_PROC & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en ejecución"
    Set rangoFiltro = Nothing
    Set tablaActiva = Nothing
    Set celdaActiva = Nothing
    Set hojaActiva = Nothing
End Sub

' -----------------------------------------------------------------------------
' Validación defensiva de libro, hoja, protección y celda activa
' -----------------------------------------------------------------------------
Private Function ValidarPrecondicionesFiltro( _
    ByRef hojaDestino As Worksheet, _
    ByRef celdaObjetivo As Range, _
    ByRef tablaObjetivo As ListObject, _
    ByRef mensajeError As String _
) As Boolean
    ValidarPrecondicionesFiltro = False
    Set tablaObjetivo = Nothing

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
        mensajeError = "La hoja '" & hojaDestino.Name & "' está protegida. Desprotégela para aplicar filtros."
        Exit Function
    End If

    If ActiveCell Is Nothing Then
        mensajeError = "Selecciona una celda dentro de los datos antes de ejecutar el filtro."
        Exit Function
    End If

    Set celdaObjetivo = ActiveCell

    If IsError(celdaObjetivo.Value) Then
        mensajeError = "La celda activa (" & celdaObjetivo.Address(False, False) & ") contiene un valor de error."
        Exit Function
    End If

    ' Verificar si la celda pertenece a una tabla estructurada de Excel
    On Error Resume Next
    Set tablaObjetivo = celdaObjetivo.ListObject
    On Error GoTo 0

    ValidarPrecondicionesFiltro = True
End Function

' -----------------------------------------------------------------------------
' Determina el rango de filtrado en una hoja normal, contemplando columnas
' iniciales vacías y asegurando que la columna de la celda activa esté incluida.
' -----------------------------------------------------------------------------
Private Function ResolverRangoHojaConColumnasVacias( _
    ByVal hoja As Worksheet, _
    ByVal celdaActiva As Range, _
    ByRef rangoResultado As Range, _
    ByRef campoRelativo As Long, _
    ByRef mensajeError As String _
) As Boolean
    Dim primeraFila As Long
    Dim ultimaFila As Long
    Dim primeraColumna As Long
    Dim ultimaColumna As Long

    ResolverRangoHojaConColumnasVacias = False

    If Not ObtenerLimitesDatosHoja(hoja, primeraFila, ultimaFila, primeraColumna, ultimaColumna) Then
        mensajeError = "La hoja '" & hoja.Name & "' no contiene datos para filtrar."
        Exit Function
    End If

    ' Garantizar que la celda activa esté dentro de los límites horizontales y verticales
    If celdaActiva.Column < primeraColumna Then
        primeraColumna = celdaActiva.Column
    End If

    If celdaActiva.Column > ultimaColumna Then
        ultimaColumna = celdaActiva.Column
    End If

    If celdaActiva.Row > ultimaFila Then
        ultimaFila = celdaActiva.Row
    End If

    ' Si la celda activa está en la fila 1 o por encima del inicio de datos, iniciar en primeraFila
    If primeraFila < 1 Then
        primeraFila = 1
    End If

    ' Debe haber al menos encabezado y una fila de datos
    If ultimaFila <= primeraFila Then
        mensajeError = "Se requiere al menos una fila de encabezado y una fila de datos para filtrar."
        Exit Function
    End If

    Set rangoResultado = hoja.Range( _
        hoja.Cells(primeraFila, primeraColumna), _
        hoja.Cells(ultimaFila, ultimaColumna) _
    )

    ' El índice de campo en AutoFilter es relativo a la primera columna del rango filtrado
    campoRelativo = celdaActiva.Column - primeraColumna + 1

    If campoRelativo < 1 Or campoRelativo > rangoResultado.Columns.Count Then
        mensajeError = "No se pudo determinar la columna relativa para el filtro."
        Exit Function
    End If

    ResolverRangoHojaConColumnasVacias = True
End Function

' -----------------------------------------------------------------------------
' Obtiene primera/última fila y primera/última columna con contenido real
' -----------------------------------------------------------------------------
Public Function ObtenerLimitesDatosHoja( _
    ByVal hoja As Worksheet, _
    ByRef primeraFila As Long, _
    ByRef ultimaFila As Long, _
    ByRef primeraColumna As Long, _
    ByRef ultimaColumna As Long _
) As Boolean
    Dim celdaMinRow As Range
    Dim celdaMaxRow As Range
    Dim celdaMinCol As Range
    Dim celdaMaxCol As Range

    ObtenerLimitesDatosHoja = False

    On Error Resume Next
    Set celdaMinRow = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(hoja.Rows.Count, hoja.Columns.Count), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )
    Set celdaMaxRow = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )
    Set celdaMinCol = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(hoja.Rows.Count, hoja.Columns.Count), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False _
    )
    Set celdaMaxCol = hoja.Cells.Find( _
        What:="*", _
        After:=hoja.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )
    On Error GoTo 0

    If celdaMinRow Is Nothing Or celdaMaxRow Is Nothing Or celdaMinCol Is Nothing Or celdaMaxCol Is Nothing Then
        Exit Function
    End If

    primeraFila = celdaMinRow.Row
    ultimaFila = celdaMaxRow.Row
    primeraColumna = celdaMinCol.Column
    ultimaColumna = celdaMaxCol.Column

    ObtenerLimitesDatosHoja = (primeraFila >= 1 And ultimaFila >= primeraFila And _
                               primeraColumna >= 1 And ultimaColumna >= primeraColumna)
End Function

' -----------------------------------------------------------------------------
' Aplica el AutoFilter asegurando que el rango tenga AutoFilter activo
' -----------------------------------------------------------------------------
Private Sub AplicarAutoFiltroEnRango( _
    ByVal hoja As Worksheet, _
    ByVal rangoObjetivo As Range, _
    ByVal campoRelativo As Long, _
    ByVal celdaActiva As Range _
)
    Dim criterio As Variant

    ' Si existe un AutoFilter previo en la hoja que no coincide con el rango objetivo, limpiarlo
    If hoja.AutoFilterMode Then
        If hoja.AutoFilter.Range.Address <> rangoObjetivo.Address Then
            hoja.AutoFilterMode = False
        End If
    End If

    ' Activar AutoFilter sobre el rango exacto si aún no está habilitado
    If Not hoja.AutoFilterMode Then
        rangoObjetivo.AutoFilter
    End If

    ' Preparar criterio según si la celda activa está vacía o tiene valor
    If IsEmpty(celdaActiva.Value) Or Trim$(CStr(celdaActiva.Value)) = vbNullString Then
        criterio = "="
    ElseIf IsDate(celdaActiva.Value) And Not IsNumeric(celdaActiva.Value) Then
        criterio = "=" & CStr(CDate(celdaActiva.Value))
    Else
        criterio = "=" & CStr(celdaActiva.Value)
    End If

    rangoObjetivo.AutoFilter Field:=campoRelativo, Criteria1:=criterio
End Sub



