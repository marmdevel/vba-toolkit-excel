Attribute VB_Name = "modExportarVisiblesANuevaHoja"
Option Explicit

' =============================================================================
' Módulo: modExportarVisiblesANuevaHoja
' Propósito: Copia únicamente las celdas visibles de la hoja activa (filas
'            filtradas, filas ocultas/alto 0 y columnas ocultas excluidas)
'            hacia una nueva hoja dentro del libro activo.
' Entorno:   Excel 2019 (64 bits / 32 bits compatible).
' =============================================================================

' -----------------------------------------------------------------------------
' Procedimiento principal de entrada
' -----------------------------------------------------------------------------
Public Sub CopiarCeldasVisiblesANuevaHoja(Optional ByVal soloValores As Boolean = True)
    Const NOMBRE_PROC As String = "CopiarCeldasVisiblesANuevaHoja"

    Dim libroActivo As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaNueva As Worksheet
    Dim rangoVisibles As Range
    Dim estadoPrevio As EstadoEntornoExcel
    Dim mensajeValidacion As String

    On Error GoTo ManejarError

    ' 1. Validar precondiciones antes de alterar el entorno o crear hojas
    If Not ValidarPrecondicionesVisibles(libroActivo, hojaOrigen, rangoVisibles, mensajeValidacion) Then
        MsgBox mensajeValidacion, vbExclamation, "Validación de celdas visibles"
        GoTo Salir
    End If

    ' 2. Capturar y optimizar propiedades del entorno de Excel
    estadoPrevio = CapturarYDesactivarEntorno()

    ' 3. Crear la nueva hoja y transferir únicamente las celdas visibles
    Set hojaNueva = CrearHojaDestino(libroActivo, hojaOrigen)
    CopiarRangoVisibleEnHoja rangoVisibles, hojaNueva, soloValores

    MsgBox "Se copiaron las celdas visibles de '" & hojaOrigen.Name & "' " & _
           "a la nueva hoja '" & hojaNueva.Name & "'.", _
           vbInformation, "Exportación completada"

Salir:
    RestaurarEntorno estadoPrevio
    Set rangoVisibles = Nothing
    Set hojaNueva = Nothing
    Set hojaOrigen = Nothing
    Set libroActivo = Nothing
    Exit Sub

ManejarError:
    RestaurarEntorno estadoPrevio
    MsgBox "No fue posible copiar las celdas visibles a una nueva hoja." & vbCrLf & vbCrLf & _
           "Procedimiento: " & NOMBRE_PROC & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en ejecución"
    Set rangoVisibles = Nothing
    Set hojaNueva = Nothing
    Set hojaOrigen = Nothing
    Set libroActivo = Nothing
End Sub

' -----------------------------------------------------------------------------
' Validación de libro, hoja, protección y existencia de celdas visibles
' -----------------------------------------------------------------------------
Private Function ValidarPrecondicionesVisibles( _
    ByRef libroDestino As Workbook, _
    ByRef hojaFuente As Worksheet, _
    ByRef rangoVisiblesSalida As Range, _
    ByRef mensajeError As String _
) As Boolean
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoDatos As Range

    ValidarPrecondicionesVisibles = False
    Set rangoVisiblesSalida = Nothing

    If ActiveWorkbook Is Nothing Then
        mensajeError = "No hay un libro de Excel activo."
        Exit Function
    End If

    Set libroDestino = ActiveWorkbook

    If Not TypeOf ActiveSheet Is Worksheet Then
        mensajeError = "La hoja activa debe ser una hoja de cálculo válida."
        Exit Function
    End If

    Set hojaFuente = ActiveSheet

    If hojaFuente.ProtectContents Then
        mensajeError = "La hoja '" & hojaFuente.Name & "' está protegida. Desprotégela antes de copiar datos."
        Exit Function
    End If

    If Not ObtenerUltimaCeldaUsada(hojaFuente, ultimaFila, ultimaColumna) Then
        mensajeError = "La hoja '" & hojaFuente.Name & "' no contiene datos."
        Exit Function
    End If

    Set rangoDatos = hojaFuente.Range(hojaFuente.Cells(1, 1), hojaFuente.Cells(ultimaFila, ultimaColumna))

    ' Obtener solo celdas visibles (excluye filas filtradas, alto 0/ocultas y columnas ocultas)
    On Error Resume Next
    Set rangoVisiblesSalida = rangoDatos.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rangoVisiblesSalida Is Nothing Then
        mensajeError = "No se encontraron celdas visibles dentro del rango de datos de '" & hojaFuente.Name & "'."
        Exit Function
    End If

    ValidarPrecondicionesVisibles = True
End Function

' -----------------------------------------------------------------------------
' Crea una hoja nueva después de la hoja origen con nombre único válido
' -----------------------------------------------------------------------------
Private Function CrearHojaDestino( _
    ByVal libro As Workbook, _
    ByVal hojaReferencia As Worksheet _
) As Worksheet
    Dim nuevaHoja As Worksheet
    Dim nombreBase As String
    Dim nombreCandidato As String
    Dim contador As Long

    Set nuevaHoja = libro.Worksheets.Add(After:=hojaReferencia)

    ' Limitar base a máx. 24 caracteres para dejar espacio a sufijo y respetar límite de 31 de Excel
    nombreBase = Left$(hojaReferencia.Name, 20) & "_Vis"
    nombreCandidato = nombreBase
    contador = 1

    Do While ExisteNombreHoja(libro, nombreCandidato)
        contador = contador + 1
        nombreCandidato = Left$(nombreBase, 27) & "_" & CStr(contador)
    Loop

    nuevaHoja.Name = nombreCandidato
    Set CrearHojaDestino = nuevaHoja
End Function

Private Function ExisteNombreHoja( _
    ByVal libro As Workbook, _
    ByVal nombreBuscado As String _
) As Boolean
    Dim h As Worksheet

    For Each h In libro.Worksheets
        If StrComp(h.Name, nombreBuscado, vbTextCompare) = 0 Then
            ExisteNombreHoja = True
            Exit Function
        End If
    Next h

    ExisteNombreHoja = False
End Function

' -----------------------------------------------------------------------------
' Copia el rango visible a la celda A1 de la hoja destino
' -----------------------------------------------------------------------------
Private Sub CopiarRangoVisibleEnHoja( _
    ByVal rangoVisible As Range, _
    ByVal hojaDestino As Worksheet, _
    ByVal soloValores As Boolean _
)
    Dim destinoInicio As Range
    Set destinoInicio = hojaDestino.Cells(1, 1)

    rangoVisible.Copy

    If soloValores Then
        destinoInicio.PasteSpecial Paste:=xlPasteValues, _
                                 Operation:=xlNone, _
                                 SkipBlanks:=False, _
                                 Transpose:=False
    Else
        destinoInicio.PasteSpecial Paste:=xlPasteAll, _
                                 Operation:=xlNone, _
                                 SkipBlanks:=False, _
                                 Transpose:=False
    End If

    Application.CutCopyMode = False
End Sub

