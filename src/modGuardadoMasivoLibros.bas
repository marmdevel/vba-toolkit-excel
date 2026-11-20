Attribute VB_Name = "modGuardadoMasivoLibros"
Option Explicit

' =============================================================================
' Módulo: modGuardadoMasivoLibros
' Propósito: Guardar todos los libros abiertos que ya cuenten con ruta en disco
'           e informar cuáles no se guardaron por no haberse guardado antes.
' Entorno:  Excel 2019 (64 bits) / Compatible 32 y 64 bits.
' =============================================================================

Private Const NOMBRE_MODULO As String = "modGuardadoMasivoLibros"

' -----------------------------------------------------------------------------
' Procedimiento principal: Orquesta el guardado seguro de libros ya guardados
' y muestra el resultado al usuario.
' -----------------------------------------------------------------------------
Public Sub GuardarLibrosAbiertos()
    Const NOMBRE_PROC As String = "GuardarLibrosAbiertos"

    ' Guardar propiedades de entorno para restaurarlas siempre en Salir:
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalcMode As XlCalculation
    Dim estadoGuardado As Boolean

    Dim librosNoGuardados As Collection
    Dim totalGuardados As Long

    On Error GoTo ManejarError

    ' Validación previa: verificar que existan libros abiertos
    If Application.Workbooks.Count = 0 Then
        MsgBox "No hay libros abiertos en Excel para guardar.", _
               vbExclamation, NOMBRE_MODULO & "." & NOMBRE_PROC
        GoTo Salir
    End If

    ' Optimización de entorno durante el proceso
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents
    prevDisplayAlerts = Application.displayAlerts
    prevCalcMode = Application.calculation
    estadoGuardado = True

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.displayAlerts = False
    Application.calculation = xlCalculationManual

    Set librosNoGuardados = New Collection

    ' Ejecutar lógica de guardado y recolección de omitidos
    ProcesarGuardadoLibros librosNoGuardados, totalGuardados

    ' Informar el resultado al usuario según los libros omitidos
    MostrarResumenGuardado totalGuardados, librosNoGuardados

Salir:
    On Error Resume Next
    If estadoGuardado Then
        Application.calculation = prevCalcMode
        Application.displayAlerts = prevDisplayAlerts
        Application.enableEvents = prevEnableEvents
        Application.screenUpdating = prevScreenUpdating
    End If
    Set librosNoGuardados = Nothing
    On Error GoTo 0
    Exit Sub

ManejarError:
    MsgBox "Ocurrió un error al intentar guardar los libros abiertos." & vbCrLf & _
           "Procedimiento: " & NOMBRE_MODULO & "." & NOMBRE_PROC & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en " & NOMBRE_MODULO
    GoTo Salir
End Sub

' -----------------------------------------------------------------------------
' Recorre los libros abiertos, guarda los que ya tienen ruta en disco y registra
' los que nunca han sido guardados previamente.
' -----------------------------------------------------------------------------
Private Sub ProcesarGuardadoLibros( _
    ByRef noGuardados As Collection, _
    ByRef outTotalGuardados As Long _
)
    Dim wbActual As Workbook
    Dim idx As Long

    outTotalGuardados = 0

    ' Se recorre por índice para mayor estabilidad si cambia la colección
    For idx = 1 To Application.Workbooks.Count
        Set wbActual = Application.Workbooks(idx)

        If EsLibroGuardadoPreviamente(wbActual) Then
            ' Evitar guardar complementos (AddIns) no deseados si aplica;
            ' aquí se guardan libros normales abiertos.
            If Not wbActual.IsAddin Then
                wbActual.Save
                outTotalGuardados = outTotalGuardados + 1
            End If
        Else
            noGuardados.Add wbActual.Name
        End If
    Next idx

    Set wbActual = Nothing
End Sub

' -----------------------------------------------------------------------------
' Determina si el libro ya existe en disco evaluando su propiedad Path.
' Un libro nuevo (ej. "Libro1") tiene Path vacío hasta su primer guardado.
' -----------------------------------------------------------------------------
Private Function EsLibroGuardadoPreviamente(ByVal wbObjetivo As Workbook) As Boolean
    If wbObjetivo Is Nothing Then
        EsLibroGuardadoPreviamente = False
        Exit Function
    End If

    EsLibroGuardadoPreviamente = (LenB(wbObjetivo.Path) > 0)
End Function

' -----------------------------------------------------------------------------
' Construye y muestra el mensaje final con los libros guardados y los que no
' se guardaron por no haber sido guardados previamente.
' -----------------------------------------------------------------------------
Private Sub MostrarResumenGuardado( _
    ByVal totalGuardados As Long, _
    ByVal noGuardados As Collection _
)
    Dim mensaje As String
    Dim itemNombre As Variant
    Dim detalleNoGuardados As String

    If noGuardados Is Nothing Or noGuardados.Count = 0 Then
        mensaje = "Proceso completado. Se guardaron " & totalGuardados & " libro(s)." & vbCrLf & _
                  "Todos los libros abiertos ya contaban con guardado previo."
        MsgBox mensaje, vbInformation, "Guardado completado"
        Exit Sub
    End If

    For Each itemNombre In noGuardados
        detalleNoGuardados = detalleNoGuardados & " - " & CStr(itemNombre) & vbCrLf
    Next itemNombre

    mensaje = "Proceso finalizado con avisos." & vbCrLf & _
              "Libros guardados: " & totalGuardados & vbCrLf & vbCrLf & _
              "No se guardaron los siguientes libros por no haber sido guardados previamente:" & vbCrLf & _
              detalleNoGuardados

    MsgBox mensaje, vbExclamation, "Libros pendientes de primer guardado"
End Sub

