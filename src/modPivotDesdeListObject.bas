Attribute VB_Name = "modPivotDesdeListObject"
Option Explicit

' =============================================================================
' Módulo: modPivotDesdeListObject
' Propósito: Crear una Tabla Dinámica vacía/base en una nueva hoja a partir
'           del ListObject que contiene una celda específica.
' Entorno:  Excel 2019 (64 bits) / Compatible 32 y 64 bits.
' =============================================================================

Private Const NOMBRE_MODULO As String = "modPivotDesdeListObject"

' -----------------------------------------------------------------------------
' Función principal: Valida que la celda esté dentro de un ListObject y crea
' una PivotTable en una hoja nueva. Devuelve el PivotTable creado o Nothing.
' -----------------------------------------------------------------------------
Public Function CrearTablaDinamicaDesdeCelda( _
    ByVal celdaReferencia As Range, _
    Optional ByVal nombreHojaDestino As String = vbNullString, _
    Optional ByVal mostrarMensajes As Boolean = True _
) As PivotTable
    Const NOMBRE_PROC As String = "CrearTablaDinamicaDesdeCelda"

    ' Guardar propiedades de entorno para restaurarlas siempre en Salir:
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalcMode As XlCalculation
    Dim estadoGuardado As Boolean

    Dim wsOrigen As Worksheet
    Dim wbOrigen As Workbook
    Dim loTabla As ListObject
    Dim wsNueva As Worksheet
    Dim pcCache As PivotCache
    Dim ptCreada As PivotTable
    Dim nombreTablaPivot As String

    On Error GoTo ManejarError

    Set CrearTablaDinamicaDesdeCelda = Nothing

    ' 1. Validaciones previas (programación defensiva)
    If Not ValidarCeldaEnListObject(celdaReferencia, wsOrigen, wbOrigen, loTabla, mostrarMensajes) Then
        GoTo Salir
    End If

    ' 2. Guardar y optimizar estado de Excel
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents
    prevDisplayAlerts = Application.displayAlerts
    prevCalcMode = Application.calculation
    estadoGuardado = True

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.displayAlerts = False
    Application.calculation = xlCalculationManual

    ' 3. Crear hoja nueva para la tabla dinámica
    Set wsNueva = CrearHojaNueva(wbOrigen, nombreHojaDestino)

    ' 4. Crear PivotCache y PivotTable sobre el rango completo del ListObject
    nombreTablaPivot = GenerarNombrePivotUnico(wsNueva, loTabla.Name)

    Set pcCache = wbOrigen.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=loTabla.Range, _
        Version:=xlPivotTableVersion15 _
    )

    Set ptCreada = pcCache.CreatePivotTable( _
        TableDestination:=wsNueva.Range("A3"), _
        TableName:=nombreTablaPivot, _
        DefaultVersion:=xlPivotTableVersion15 _
    )

    ' Dejar lista para que el llamador configure campos y agrupaciones
    ptCreada.ManualUpdate = False
    Set CrearTablaDinamicaDesdeCelda = ptCreada

Salir:
    On Error Resume Next
    If estadoGuardado Then
        Application.calculation = prevCalcMode
        Application.displayAlerts = prevDisplayAlerts
        Application.enableEvents = prevEnableEvents
        Application.screenUpdating = prevScreenUpdating
    End If
    Set ptCreada = Nothing
    Set pcCache = Nothing
    Set wsNueva = Nothing
    Set loTabla = Nothing
    Set wsOrigen = Nothing
    Set wbOrigen = Nothing
    On Error GoTo 0
    Exit Function

ManejarError:
    Set CrearTablaDinamicaDesdeCelda = Nothing

    If mostrarMensajes Then
        MsgBox "Ocurrió un error al crear la tabla dinámica." & vbCrLf & _
               "Procedimiento: " & NOMBRE_MODULO & "." & NOMBRE_PROC & vbCrLf & _
               "Número de error: " & Err.Number & vbCrLf & _
               "Descripción: " & Err.Description, _
               vbCritical, "Error en " & NOMBRE_MODULO
    End If
    GoTo Salir
End Function

' -----------------------------------------------------------------------------
' Valida que la celda sea válida, pertenezca a una hoja y esté contenida
' dentro de un ListObject con filas de datos.
' -----------------------------------------------------------------------------
Private Function ValidarCeldaEnListObject( _
    ByVal celdaRef As Range, _
    ByRef outWs As Worksheet, _
    ByRef outWb As Workbook, _
    ByRef outTabla As ListObject, _
    ByVal mostrarMensajes As Boolean _
) As Boolean
    Set outWs = Nothing
    Set outWb = Nothing
    Set outTabla = Nothing
    ValidarCeldaEnListObject = False

    If celdaRef Is Nothing Then
        If mostrarMensajes Then
            MsgBox "Debe indicar una celda de referencia válida.", vbExclamation, NOMBRE_MODULO
        End If
        Exit Function
    End If

    If TypeName(celdaRef.Parent) <> "Worksheet" Then
        If mostrarMensajes Then
            MsgBox "La celda debe pertenecer a una hoja de cálculo válida.", vbExclamation, NOMBRE_MODULO
        End If
        Exit Function
    End If

    Set outWs = celdaRef.Parent
    Set outWb = outWs.Parent

    ' Identifica la tabla específica en la que se encuentra la celda
    Set outTabla = celdaRef.Cells(1, 1).ListObject

    If outTabla Is Nothing Then
        If mostrarMensajes Then
            MsgBox "La celda '" & celdaRef.Cells(1, 1).Address(False, False) & _
                   "' en la hoja '" & outWs.Name & "' no pertenece a ninguna tabla (ListObject).", _
                   vbExclamation, NOMBRE_MODULO
        End If
        Exit Function
    End If

    If outTabla.DataBodyRange Is Nothing Then
        If mostrarMensajes Then
            MsgBox "La tabla '" & outTabla.Name & "' no contiene filas de datos.", _
                   vbExclamation, NOMBRE_MODULO
        End If
        Exit Function
    End If

    ValidarCeldaEnListObject = True
End Function

' -----------------------------------------------------------------------------
' Crea una hoja nueva al final del libro y le asigna un nombre válido/único.
' -----------------------------------------------------------------------------
Private Function CrearHojaNueva( _
    ByVal wb As Workbook, _
    ByVal nombreSugerido As String _
) As Worksheet
    Dim wsNueva As Worksheet
    Dim nombreLimpio As String
    Dim nombreFinal As String
    Dim contador As Long

    Set wsNueva = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))

    nombreLimpio = Left$(Trim$(nombreSugerido), 25)
    If LenB(nombreLimpio) = 0 Then
        nombreLimpio = "TD_" & Format$(Now, "yyyymmdd_hhmmss")
    End If

    nombreFinal = nombreLimpio
    contador = 1

    Do While ExisteNombreHoja(wb, nombreFinal)
        contador = contador + 1
        nombreFinal = Left$(nombreLimpio, 27) & "_" & CStr(contador)
    Loop

    wsNueva.Name = nombreFinal
    Set CrearHojaNueva = wsNueva
End Function

Private Function ExisteNombreHoja(ByVal wb As Workbook, ByVal nombreHoja As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nombreHoja)
    ExisteNombreHoja = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Function GenerarNombrePivotUnico( _
    ByVal wsDestino As Worksheet, _
    ByVal nombreBaseTabla As String _
) As String
    Dim limpio As String
    Dim candidato As String
    Dim idx As Long

    limpio = Replace(nombreBaseTabla, " ", "_")
    limpio = Left$(limpio, 20)
    If LenB(limpio) = 0 Then limpio = "Tabla"

    candidato = "PT_" & limpio & "_1"
    idx = 1

    Do While ExistePivotEnHoja(wsDestino, candidato)
        idx = idx + 1
        candidato = "PT_" & limpio & "_" & CStr(idx)
    Loop

    GenerarNombrePivotUnico = candidato
End Function

Private Function ExistePivotEnHoja(ByVal ws As Worksheet, ByVal nombrePivot As String) As Boolean
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        If StrComp(pt.Name, nombrePivot, vbTextCompare) = 0 Then
            ExistePivotEnHoja = True
            Exit Function
        End If
    Next pt
    ExistePivotEnHoja = False
End Function

' =============================================================================
' Ejemplo de uso: crea la tabla dinámica desde la celda activa y devuelve el PivotTable
' para que luego tú agregues campos o agrupes.
' =============================================================================
Public Sub EjemploCrearPivotDesdeCeldaActiva()
    Const NOMBRE_PROC As String = "EjemploCrearPivotDesdeCeldaActiva"

    Dim ptResultado As PivotTable

    On Error GoTo ManejarError

    If TypeName(ActiveSheet) <> "Worksheet" Or TypeName(ActiveCell) <> "Range" Then
        MsgBox "Seleccione una celda dentro de una tabla en una hoja de cálculo.", _
               vbExclamation, NOMBRE_PROC
        GoTo Salir
    End If

    Set ptResultado = CrearTablaDinamicaDesdeCelda( _
        celdaReferencia:=ActiveCell, _
        nombreHojaDestino:="Pivot_Generada", _
        mostrarMensajes:=True _
    )

    If Not ptResultado Is Nothing Then
        MsgBox "Tabla dinámica '" & ptResultado.Name & "' creada en la hoja '" & _
               ptResultado.Parent.Name & "'.", _
               vbInformation, "Listo para agrupar"
    End If

Salir:
    Set ptResultado = Nothing
    Exit Sub

ManejarError:
    MsgBox "Error en " & NOMBRE_PROC & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error"
    GoTo Salir
End Sub

