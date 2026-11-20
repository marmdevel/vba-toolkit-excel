Attribute VB_Name = "modCuartilesColorLeyendaHtml"
Option Explicit

' =============================================================================
' Módulo: modCuartilesColorLeyendaHtml
' Propósito: Colorear un rango seleccionado en 4 colores según cuartiles
'           (numérico por valor; alfanumérico por frecuencia de ocurrencia) y
'           generar una leyenda en HTML en una ruta fija.
' Entorno:  Excel 2019 (64 bits) / Compatible 32 y 64 bits.
' =============================================================================

Private Const NOMBRE_MODULO As String = "modCuartilesColorLeyendaHtml"

' Ruta fija (hardcodeada) donde se generará la página HTML de leyenda
Private Const RUTA_LEYENDA_HTML As String = "C:\Users\PROPIETARIO\OneDrive\Documentos\Antonio\Proyectos\ToolExcel\LeyendaCuartiles.html"

' Paleta de 4 colores (RGB) para Q1, Q2, Q3 y Q4
Private Const COLOR_Q1_HEX As String = "#D9EAD3"
Private Const COLOR_Q2_HEX As String = "#FFF2CC"
Private Const COLOR_Q3_HEX As String = "#FCE5CD"
Private Const COLOR_Q4_HEX As String = "#F4CCCC"

Private Enum ModoRango
    ModoIndefinido = 0
    ModoNumerico = 1
    ModoAlfanumerico = 2
End Enum

Private Type InfoCuartiles
    modo As ModoRango
    CantidadDatos As Long
    MinVal As Double
    Q1Val As Double
    MedianaVal As Double
    Q3Val As Double
    MaxVal As Double
End Type

' -----------------------------------------------------------------------------
' Procedimiento principal: Valida la selección, clasifica por cuartiles,
' colorea las celdas y genera el archivo HTML de leyenda.
' -----------------------------------------------------------------------------
Public Sub ColorearRangoPorCuartilesYGenerarHtml()
    Const NOMBRE_PROC As String = "ColorearRangoPorCuartilesYGenerarHtml"

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalcMode As XlCalculation
    Dim estadoGuardado As Boolean

    Dim rngSel As Range
    Dim wsActiva As Worksheet
    Dim datos As Variant
    Dim frecuencias As Object
    Dim valoresBase() As Double
    Dim infoQ As InfoCuartiles
    Dim esSoloNumerico As Boolean

    On Error GoTo ManejarError

    ' 1. Validaciones previas defensivas
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "La vista activa no es una hoja de cálculo válida.", vbExclamation, NOMBRE_MODULO
        GoTo Salir
    End If

    Set wsActiva = ActiveSheet

    If wsActiva.ProtectContents Then
        MsgBox "La hoja '" & wsActiva.Name & "' está protegida. Desprotéjala antes de aplicar colores.", _
               vbExclamation, NOMBRE_MODULO
        GoTo Salir
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Debe seleccionar un rango de celdas válido.", vbExclamation, NOMBRE_MODULO
        GoTo Salir
    End If

    Set rngSel = Selection

    If rngSel.Areas.Count > 1 Then
        MsgBox "Seleccione un único rango contiguo (no selecciones múltiples áreas separadas).", _
               vbExclamation, NOMBRE_MODULO
        GoTo Salir
    End If

    If rngSel.CountLarge = 0 Then
        MsgBox "La selección está vacía.", vbExclamation, NOMBRE_MODULO
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

    ' 3. Analizar contenido en memoria (evita bucles directos repetitivos de lectura)
    datos = ObtenerArrayRango(rngSel)
    Set frecuencias = CreateObject("Scripting.Dictionary")
    frecuencias.CompareMode = vbTextCompare

    If Not AnalizarRangoSeleccionado(datos, esSoloNumerico, valoresBase, frecuencias) Then
        MsgBox "El rango seleccionado no contiene celdas válidas (o contiene valores de error como #REF!/#N/A).", _
               vbExclamation, NOMBRE_MODULO
        GoTo Salir
    End If

    ' 4. Calcular cuartiles según modo numérico o alfanumérico
    If esSoloNumerico Then
        infoQ = CalcularCuartilesDesdeArray(valoresBase, ModoNumerico)
    Else
        infoQ = CalcularCuartilesDesdeFrecuencias(frecuencias, ModoAlfanumerico)
    End If

    ' 5. Aplicar colores de los 4 cuartiles a las celdas del rango
    PintarRangoPorCuartiles rngSel, datos, frecuencias, infoQ

    ' 6. Generar archivo HTML de leyenda en la ruta fija
    ExportarLeyendaHtml RUTA_LEYENDA_HTML, wsActiva.Name, rngSel.Address(False, False), infoQ

    MsgBox "Proceso completado correctamente." & vbCrLf & _
           "Modo aplicado: " & IIf(infoQ.modo = ModoNumerico, "Numérico (por valor)", "Alfanumérico (por ocurrencias)") & vbCrLf & _
           "Leyenda HTML generada en: " & RUTA_LEYENDA_HTML, _
           vbInformation, "Cuartiles y Leyenda HTML"

Salir:
    On Error Resume Next
    If estadoGuardado Then
        Application.calculation = prevCalcMode
        Application.displayAlerts = prevDisplayAlerts
        Application.enableEvents = prevEnableEvents
        Application.screenUpdating = prevScreenUpdating
    End If
    Set frecuencias = Nothing
    Set rngSel = Nothing
    Set wsActiva = Nothing
    On Error GoTo 0
    Exit Sub

ManejarError:
    MsgBox "Ocurrió un error al procesar los cuartiles o generar la leyenda HTML." & vbCrLf & _
           "Procedimiento: " & NOMBRE_MODULO & "." & NOMBRE_PROC & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en " & NOMBRE_MODULO
    GoTo Salir
End Sub


' -----------------------------------------------------------------------------
' Normaliza el contenido de un Range a un array bidimensional (1..N, 1..M).
' -----------------------------------------------------------------------------
Private Function ObtenerArrayRango(ByVal rngObjetivo As Range) As Variant
    Dim arrTmp As Variant
    Dim arrNormalizado(1 To 1, 1 To 1) As Variant

    If rngObjetivo.CountLarge = 1 Then
        arrNormalizado(1, 1) = rngObjetivo.Value
        ObtenerArrayRango = arrNormalizado
        Exit Function
    End If

    arrTmp = rngObjetivo.Value
    If IsArray(arrTmp) Then
        ObtenerArrayRango = arrTmp
    Else
        arrNormalizado(1, 1) = arrTmp
        ObtenerArrayRango = arrNormalizado
    End If
End Function

' -----------------------------------------------------------------------------
' Examina el array del rango y determina si todos los datos no vacíos son
' numéricos o si debe tratarse como alfanumérico (contando ocurrencias).
' -----------------------------------------------------------------------------
Private Function AnalizarRangoSeleccionado( _
    ByRef datos As Variant, _
    ByRef outEsSoloNumerico As Boolean, _
    ByRef outValoresNum() As Double, _
    ByRef outFrecuencias As Object _
) As Boolean
    Dim r As Long, c As Long
    Dim maxR As Long, maxC As Long
    Dim valCelda As Variant
    Dim textoClave As String
    Dim countNum As Long

    maxR = UBound(datos, 1)
    maxC = UBound(datos, 2)
    outEsSoloNumerico = True
    countNum = 0

    ' Primera pasada: verificar errores y si todos los valores no vacíos son numéricos
    For r = 1 To maxR
        For c = 1 To maxC
            valCelda = datos(r, c)
            If Not IsEmpty(valCelda) Then
                If IsError(valCelda) Then
                    AnalizarRangoSeleccionado = False
                    Exit Function
                End If

                If LenB(Trim$(CStr(valCelda))) > 0 Then
                    If VarType(valCelda) = vbBoolean Or Not IsNumeric(valCelda) Then
                        outEsSoloNumerico = False
                    End If
                End If
            End If
        Next c
    Next r

    ' Segunda pasada: cargar array numérico o diccionario de frecuencias
    If outEsSoloNumerico Then
        For r = 1 To maxR
            For c = 1 To maxC
                valCelda = datos(r, c)
                If Not IsEmpty(valCelda) Then
                    If LenB(Trim$(CStr(valCelda))) > 0 Then
                        If countNum = 0 Then
                            ReDim outValoresNum(1 To 1)
                        Else
                            ReDim Preserve outValoresNum(1 To countNum + 1)
                        End If
                        countNum = countNum + 1
                        outValoresNum(countNum) = CDbl(valCelda)
                    End If
                End If
            Next c
        Next r
        AnalizarRangoSeleccionado = (countNum > 0)
    Else
        For r = 1 To maxR
            For c = 1 To maxC
                valCelda = datos(r, c)
                If Not IsEmpty(valCelda) Then
                    textoClave = Trim$(CStr(valCelda))
                    If LenB(textoClave) > 0 Then
                        If outFrecuencias.Exists(textoClave) Then
                            outFrecuencias(textoClave) = CLng(outFrecuencias(textoClave)) + 1
                        Else
                            outFrecuencias.Add textoClave, 1&
                        End If
                    End If
                End If
            Next c
        Next r
        AnalizarRangoSeleccionado = (outFrecuencias.Count > 0)
    End If
End Function

' -----------------------------------------------------------------------------
' Calcula cuartiles a partir de un array numérico.
' -----------------------------------------------------------------------------
Private Function CalcularCuartilesDesdeArray( _
    ByRef valores() As Double, _
    ByVal modo As ModoRango _
) As InfoCuartiles
    Dim info As InfoCuartiles
    Dim i As Long
    Dim minV As Double, maxV As Double

    info.modo = modo
    info.CantidadDatos = UBound(valores)

    minV = valores(1)
    maxV = valores(1)
    For i = 2 To UBound(valores)
        If valores(i) < minV Then minV = valores(i)
        If valores(i) > maxV Then maxV = valores(i)
    Next i

    info.MinVal = minV
    info.MaxVal = maxV
    info.Q1Val = CDbl(Application.WorksheetFunction.Quartile(valores, 1))
    info.MedianaVal = CDbl(Application.WorksheetFunction.Quartile(valores, 2))
    info.Q3Val = CDbl(Application.WorksheetFunction.Quartile(valores, 3))

    CalcularCuartilesDesdeArray = info
End Function

' -----------------------------------------------------------------------------
' Calcula cuartiles a partir de las ocurrencias de cadenas alfanuméricas.
' -----------------------------------------------------------------------------
Private Function CalcularCuartilesDesdeFrecuencias( _
    ByVal dictFreq As Object, _
    ByVal modo As ModoRango _
) As InfoCuartiles
    Dim arrFrec() As Double
    Dim claves As Variant
    Dim i As Long, n As Long

    n = dictFreq.Count
    ReDim arrFrec(1 To n)
    claves = dictFreq.Keys

    For i = 0 To n - 1
        arrFrec(i + 1) = CDbl(dictFreq(claves(i)))
    Next i

    CalcularCuartilesDesdeFrecuencias = CalcularCuartilesDesdeArray(arrFrec, modo)
End Function

Private Function FilaHtml(ByVal etiqueta As String, ByVal hexCol As String, ByVal rangoTexto As String) As String
    FilaHtml = "      <tr>" & _
               "<td><strong>" & EscapeHtml(etiqueta) & "</strong></td>" & _
               "<td><span class=""swatch"" style=""background-color:" & hexCol & ";""></span></td>" & _
               "<td>" & hexCol & "</td>" & _
               "<td>" & EscapeHtml(rangoTexto) & "</td>" & _
               "</tr>" & vbCrLf
End Function

Private Function EscapeHtml(ByVal texto As String) As String
    Dim res As String
    res = Replace(texto, "&", "&amp;")
    res = Replace(res, "<", "&lt;")
    res = Replace(res, ">", "&gt;")
    res = Replace(res, """", "&quot;")
    EscapeHtml = res
End Function

Private Sub AsegurarCarpetaExiste(ByVal rutaCarpeta As String)
    Dim fso As Object
    If LenB(rutaCarpeta) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rutaCarpeta) Then
        fso.CreateFolder rutaCarpeta
    End If
    Set fso = Nothing
End Sub

' -----------------------------------------------------------------------------
' Aplica color de fondo a cada celda del rango según su cuartil.
' -----------------------------------------------------------------------------
Private Sub PintarRangoPorCuartiles( _
    ByVal rngObjetivo As Range, _
    ByRef datos As Variant, _
    ByVal dictFreq As Object, _
    ByRef infoQ As InfoCuartiles _
)
    Dim r As Long, c As Long
    Dim maxR As Long, maxC As Long
    Dim valCelda As Variant
    Dim textoClave As String
    Dim valorEvaluado As Double
    Dim hexColor As String
    Dim rColor As Long, gColor As Long, bColor As Long

    maxR = UBound(datos, 1)
    maxC = UBound(datos, 2)

    For r = 1 To maxR
        For c = 1 To maxC
            valCelda = datos(r, c)
            textoClave = vbNullString
            If Not IsEmpty(valCelda) Then textoClave = Trim$(CStr(valCelda))

            If LenB(textoClave) > 0 Then
                If infoQ.modo = ModoNumerico Then
                    If IsNumeric(valCelda) And VarType(valCelda) <> vbBoolean Then
                        valorEvaluado = CDbl(valCelda)
                        hexColor = ObtenerHexPorValor(valorEvaluado, infoQ)
                        HexToRgb hexColor, rColor, gColor, bColor
                        rngObjetivo.Cells(r, c).Interior.Color = RGB(rColor, gColor, bColor)
                    End If
                Else
                    If dictFreq.Exists(textoClave) Then
                        valorEvaluado = CDbl(dictFreq(textoClave))
                        hexColor = ObtenerHexPorValor(valorEvaluado, infoQ)
                        HexToRgb hexColor, rColor, gColor, bColor
                        rngObjetivo.Cells(r, c).Interior.Color = RGB(rColor, gColor, bColor)
                    End If
                End If
            End If
        Next c
    Next r
End Sub

' -----------------------------------------------------------------------------
' Determina el color hexadecimal según el cuartil en el que cae el valor.
' -----------------------------------------------------------------------------
Private Function ObtenerHexPorValor(ByVal valor As Double, ByRef infoQ As InfoCuartiles) As String
    If valor <= infoQ.Q1Val Then
        ObtenerHexPorValor = COLOR_Q1_HEX
    ElseIf valor <= infoQ.MedianaVal Then
        ObtenerHexPorValor = COLOR_Q2_HEX
    ElseIf valor <= infoQ.Q3Val Then
        ObtenerHexPorValor = COLOR_Q3_HEX
    Else
        ObtenerHexPorValor = COLOR_Q4_HEX
    End If
End Function

Private Sub HexToRgb(ByVal hexStr As String, ByRef outR As Long, ByRef outG As Long, ByRef outB As Long)
    Dim limpio As String
    limpio = Replace(Trim$(hexStr), "#", vbNullString)
    outR = CLng("&H" & Mid$(limpio, 1, 2))
    outG = CLng("&H" & Mid$(limpio, 3, 2))
    outB = CLng("&H" & Mid$(limpio, 5, 2))
End Sub


' -----------------------------------------------------------------------------
' Genera la página HTML con la leyenda de los 4 colores y sus intervalos.
' -----------------------------------------------------------------------------
Private Sub ExportarLeyendaHtml( _
    ByVal rutaArchivo As String, _
    ByVal nombreHoja As String, _
    ByVal direccionRango As String, _
    ByRef infoQ As InfoCuartiles _
)
    Dim carpetaDestino As String
    Dim unidadStr As String
    Dim modoTexto As String
    Dim html As String
    Dim numArchivo As Integer

    carpetaDestino = Left$(rutaArchivo, InStrRev(rutaArchivo, "\") - 1)
    AsegurarCarpetaExiste carpetaDestino

    If infoQ.modo = ModoNumerico Then
        modoTexto = "Numérico (clasificación por valor de celda)"
        unidadStr = ""
    Else
        modoTexto = "Alfanumérico (clasificación por frecuencia de ocurrencias)"
        unidadStr = " ocurrencias"
    End If
    
    html = "<!DOCTYPE html>" & vbCrLf & _
           "<html lang=""es"">" & vbCrLf & _
           "<head>" & vbCrLf & _
           "  <meta charset=""UTF-8"">" & vbCrLf & _
           "  <title>Leyenda de Cuartiles</title>" & vbCrLf & _
           "  <style>" & vbCrLf & _
           "    body { font-family: Arial, sans-serif; margin: 24px; color: #222; }" & vbCrLf & _
           "    h1 { font-size: 20px; margin-bottom: 6px; }" & vbCrLf & _
           "    .meta { color: #555; margin-bottom: 18px; font-size: 14px; }" & vbCrLf & _
           "    table { border-collapse: collapse; width: 100%; max-width: 720px; }" & vbCrLf & _
           "    th, td { border: 1px solid #ccc; padding: 10px 12px; text-align: left; }" & vbCrLf & _
           "    th { background-color: #f4f4f4; }" & vbCrLf & _
           "    .swatch { width: 48px; height: 22px; border: 1px solid #888; display: inline-block; }" & vbCrLf & _
           "  </style>" & vbCrLf & _
           "</head>" & vbCrLf & _
           "<body>" & vbCrLf
    
    html = html & vbCrLf & "  <h1>Leyenda de Colores por Cuartiles</h1>" & vbCrLf & _
           "  <div class=""meta"">" & _
           "    <strong>Hoja:</strong> " & EscapeHtml(nombreHoja) & " | " & _
           "    <strong>Rango:</strong> " & EscapeHtml(direccionRango) & " | " & _
           "    <strong>Modo:</strong> " & EscapeHtml(modoTexto) & " | " & _
           "    <strong>Elementos evaluados:</strong> " & CStr(infoQ.CantidadDatos) & _
           "  </div>" & vbCrLf & _
           "  <table>" & vbCrLf & _
           "    <thead><tr><th>Cuartil</th><th>Color</th><th>Hex</th><th>Intervalo correspondiente</th></tr></thead>" & vbCrLf & _
           "    <tbody>" & vbCrLf & _
           FilaHtml("Cuartil 1 (Q1 - Bajo)", COLOR_Q1_HEX, Format$(infoQ.MinVal, "0.##") & unidadStr & " a " & Format$(infoQ.Q1Val, "0.##") & unidadStr) & _
           FilaHtml("Cuartil 2 (Q2 - Medio-Bajo)", COLOR_Q2_HEX, "> " & Format$(infoQ.Q1Val, "0.##") & unidadStr & " hasta " & Format$(infoQ.MedianaVal, "0.##") & unidadStr) & _
           FilaHtml("Cuartil 3 (Q3 - Medio-Alto)", COLOR_Q3_HEX, "> " & Format$(infoQ.MedianaVal, "0.##") & unidadStr & " hasta " & Format$(infoQ.Q3Val, "0.##") & unidadStr) & _
           FilaHtml("Cuartil 4 (Q4 - Alto)", COLOR_Q4_HEX, "> " & Format$(infoQ.Q3Val, "0.##") & unidadStr & " hasta " & Format$(infoQ.MaxVal, "0.##") & unidadStr) & _
           "    </tbody>" & vbCrLf & _
           "  </table>" & vbCrLf & _
           "</body>" & vbCrLf & _
           "</html>"
    
    numArchivo = FreeFile
    Open rutaArchivo For Output As #numArchivo
    Print #numArchivo, html
    Close #numArchivo
End Sub
