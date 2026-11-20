Attribute VB_Name = "modNamedRangeReader"
Option Explicit

' =============================================================================
' Módulo: modNamedRangeReader
' Propósito: Obtener de forma segura el valor de una celda o el texto unido de
'           un rango definido por nombre dentro de un libro específico.
' Entorno:  Excel 2019 (64 bits) / Compatible 32 y 64 bits.
' =============================================================================

Private Const NOMBRE_MODULO As String = "modNamedRangeReader"
Private Const DELIMITADOR_RANGO As String = " "

' -----------------------------------------------------------------------------
' Función principal: Intenta obtener el valor de un nombre definido en un libro.
' Devuelve True si el nombre existe y resuelve a un rango válido; False en caso contrario.
' -----------------------------------------------------------------------------
Public Function TryGetNamedCellValue( _
    ByVal wbDestino As Workbook, _
    ByVal nombreDefinido As String, _
    ByRef outValor As Variant, _
    Optional ByVal mostrarMensajes As Boolean = True _
) As Boolean
    Const NOMBRE_PROC As String = "TryGetNamedCellValue"

    ' Guardar estado original de la aplicación para restaurarlo en Salir:
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalcMode As XlCalculation
    Dim estadoGuardado As Boolean

    Dim rangoObjetivo As Range
    Dim mensajeValidacion As String

    On Error GoTo ManejarError

    ' Valor por defecto ante cualquier salida temprana o error
    outValor = Empty
    TryGetNamedCellValue = False

    ' 1. Validaciones previas (programación defensiva)
    If Not ValidateInputs(wbDestino, nombreDefinido, mensajeValidacion) Then
        If mostrarMensajes And LenB(mensajeValidacion) > 0 Then
            MsgBox mensajeValidacion, vbExclamation, NOMBRE_MODULO & "." & NOMBRE_PROC
        End If
        GoTo Salir
    End If

    ' 2. Capturar y optimizar propiedades de entorno de Excel
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents
    prevDisplayAlerts = Application.displayAlerts
    prevCalcMode = Application.calculation
    estadoGuardado = True

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.displayAlerts = False
    Application.calculation = xlCalculationManual

    ' 3. Resolver el nombre definido hacia un objeto Range válido
    If Not ResolveNamedRange(wbDestino, Trim$(nombreDefinido), rangoObjetivo) Then
        TryGetNamedCellValue = False
        GoTo Salir
    End If

    ' 4. Extraer el valor de celda única o unir el texto si es un rango múltiple
    outValor = ExtractRangeText(rangoObjetivo)
    TryGetNamedCellValue = True

Salir:
    On Error Resume Next
    If estadoGuardado Then
        Application.calculation = prevCalcMode
        Application.displayAlerts = prevDisplayAlerts
        Application.enableEvents = prevEnableEvents
        Application.screenUpdating = prevScreenUpdating
    End If
    Set rangoObjetivo = Nothing
    On Error GoTo 0
    Exit Function

ManejarError:
    TryGetNamedCellValue = False
    outValor = Empty

    If mostrarMensajes Then
        MsgBox "Ocurrió un error inesperado al consultar el nombre definido." & vbCrLf & _
               "Procedimiento: " & NOMBRE_MODULO & "." & NOMBRE_PROC & vbCrLf & _
               "Número de error: " & Err.Number & vbCrLf & _
               "Descripción: " & Err.Description, _
               vbCritical, "Error en " & NOMBRE_MODULO
    End If
    GoTo Salir
End Function

' -----------------------------------------------------------------------------
' Valida que el libro exista y que el nombre a buscar no esté vacío.
' -----------------------------------------------------------------------------
Private Function ValidateInputs( _
    ByVal wbDestino As Workbook, _
    ByVal nombreDefinido As String, _
    ByRef outMensaje As String _
) As Boolean
    outMensaje = vbNullString

    If wbDestino Is Nothing Then
        outMensaje = "Debe proporcionar una referencia válida a un libro de Excel (Workbook)."
        ValidateInputs = False
        Exit Function
    End If

    If LenB(Trim$(wbDestino.Name)) = 0 Then
        outMensaje = "El libro de destino no se encuentra abierto o no es válido."
        ValidateInputs = False
        Exit Function
    End If

    If LenB(Trim$(nombreDefinido)) = 0 Then
        outMensaje = "Debe indicar el nombre de la celda o rango que desea buscar."
        ValidateInputs = False
        Exit Function
    End If

    ValidateInputs = True
End Function

' -----------------------------------------------------------------------------
' Busca el objeto Name en el libro (y en la hoja activa si aplica) y obtiene su Range.
' Devuelve False si el nombre no existe o no apunta a un rango válido.
' -----------------------------------------------------------------------------
Private Function ResolveNamedRange( _
    ByVal wbDestino As Workbook, _
    ByVal nombreLimpio As String, _
    ByRef outRange As Range _
) As Boolean
    Dim namDef As Name
    Dim rngRef As Range

    Set outRange = Nothing
    ResolveNamedRange = False

    On Error Resume Next
    ' Primero buscar nombre con ámbito de libro
    Set namDef = wbDestino.Names(nombreLimpio)

    ' Si no existe a nivel de libro, intentar en la hoja activa del libro destino
    If namDef Is Nothing Then
        If Not wbDestino.ActiveSheet Is Nothing Then
            If TypeName(wbDestino.ActiveSheet) = "Worksheet" Then
                Set namDef = wbDestino.ActiveSheet.Names(nombreLimpio)
            End If
        End If
    End If

    If namDef Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If

    ' Intentar obtener el Range referenciado (falla si es constante, fórmula sin rango o #REF!)
    Set rngRef = namDef.RefersToRange
    On Error GoTo 0

    If rngRef Is Nothing Then
        Exit Function
    End If

    ' Verificar que el rango pertenezca al libro esperado
    If Not rngRef.Worksheet Is Nothing Then
        If Not rngRef.Worksheet.Parent Is wbDestino Then
            Exit Function
        End If
    Else
        Exit Function
    End If

    Set outRange = rngRef
    ResolveNamedRange = True
End Function

' -----------------------------------------------------------------------------
' Obtiene el valor de una celda o une los textos de un rango usando un array en memoria.
' -----------------------------------------------------------------------------
Private Function ExtractRangeText(ByVal rngObjetivo As Range) As Variant
    Dim datos As Variant
    Dim filaIdx As Long
    Dim colIdx As Long
    Dim totalFilas As Long
    Dim totalCols As Long
    Dim textoAcumulado As String
    Dim valorCelda As Variant

    If rngObjetivo.CountLarge = 1 Then
        ExtractRangeText = rngObjetivo.Value
        Exit Function
    End If

    ' Lectura en bloque a Variant para evitar bucles celda por celda
    datos = rngObjetivo.Value

    If Not IsArray(datos) Then
        ExtractRangeText = datos
        Exit Function
    End If

    totalFilas = UBound(datos, 1)
    totalCols = UBound(datos, 2)
    textoAcumulado = vbNullString

    For filaIdx = 1 To totalFilas
        For colIdx = 1 To totalCols
            valorCelda = datos(filaIdx, colIdx)
            If Not IsError(valorCelda) Then
                If LenB(CStr(valorCelda)) > 0 Then
                    If LenB(textoAcumulado) > 0 Then
                        textoAcumulado = textoAcumulado & DELIMITADOR_RANGO & CStr(valorCelda)
                    Else
                        textoAcumulado = CStr(valorCelda)
                    End If
                End If
            End If
        Next colIdx
    Next filaIdx

    ExtractRangeText = textoAcumulado
End Function


Public Function get_valorNombre_etool( _
    ByVal nombreDefinido As String, _
    ByRef valorObtenido As Variant, _
    Optional ByVal mostrarMensajes As Boolean = True _
) As Boolean

    Const NOMBRE_PROC As String = "get_valorNombre_etool"
    
    Dim wbObjetivo As Workbook
    'Dim valorObtenido As Variant
    Dim encontrado As Boolean
    Dim nombreBuscado As String

    On Error GoTo ManejarError

    ' Ejemplo: buscar en el libro activo (o reemplazar por Workbooks("Datos.xlsx"))
    Set wbObjetivo = Workbooks(NOMBRE_LIBRO_DATOS)
    nombreBuscado = nombreDefinido

    encontrado = TryGetNamedCellValue(wbObjetivo, nombreBuscado, valorObtenido, mostrarMensajes:=True)

    If encontrado And mostrarMensajes Then
        MsgBox "Nombre encontrado: " & nombreBuscado & vbCrLf & _
               "Valor / Texto: " & CStr(valorObtenido), _
               vbInformation, "Resultado"
    Else
        If mostrarMensajes Then
            MsgBox "No se encontró el nombre '" & nombreBuscado & "' en el libro '" & wbObjetivo.Name & "'.", _
                   vbExclamation, "No encontrado"
        End If
    End If

Salir:
    Set wbObjetivo = Nothing
    get_valorNombre_etool = True
    Exit Function

ManejarError:
    MsgBox "Error en " & NOMBRE_MODULO & "." & NOMBRE_PROC & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error"
    GoTo Salir
    
    
End Function

' =============================================================================
' función para invocacion
' =============================================================================
Public Function EjemploConsultarNombre()

Dim lb_encontro  As Boolean
Dim outValor As Variant


lb_encontro = get_valorNombre_etool("CW_SPACE", outValor, False)
If lb_encontro Then MsgBox CStr(outValor)

End Function
Function abrirRecursoConstante(as_recurso)

Dim lb_encontro As Boolean
Dim outValor As Variant

lb_encontro = get_valorNombre_etool(as_recurso, outValor, False)
If lb_encontro Then
    AbrirRecurso CStr(outValor)
End If
End Function
