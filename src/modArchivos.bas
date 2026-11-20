Attribute VB_Name = "modArchivos"
Option Explicit

' =========================================================================
' FUNCIÓN PARA CREAR Y GUARDAR UN NUEVO LIBRO DE EXCEL EN UNA RUTA ESPECÍFICA
' =========================================================================
' Parámetros:
' - rutaCarpeta: La ruta donde se guardará (ej: "C:\MisCarpetas\Reportes\")
' - nombreArchivo: El nombre del archivo sin extensión (ej: "ReporteDiario")
' Devuelve: True si se guardó correctamente, False si hubo error.
' =========================================================================
Public Function CrearYGuardarLibro(ByVal NombreArchivo As String, ByVal rutaCarpeta As String) As Boolean
    Dim nuevoWb As Workbook
    Dim rutaCompleta As String
    
    On Error GoTo ErrorHandler
    
    ' 1. Asegurar que la ruta termine con barra diagonal inversa (\)
    If Len(Trim(rutaCarpeta)) = 0 Then rutaCarpeta = mod_folderPrincipal
    If Len(Trim(NombreArchivo)) = 0 Then NombreArchivo = GenerarCodigoUnico
    
    
    If Right(rutaCarpeta, 1) <> "\" Then
        rutaCarpeta = rutaCarpeta & "\"
    End If
    
    ' 2. Verificar si la carpeta especificada realmente existe en el disco
    If Dir(rutaCarpeta, vbDirectory) = "" Then
        MsgBox "La ruta especificada no existe:" & vbCrLf & rutaCarpeta, vbCritical, "Ruta no encontrada"
        CrearYGuardarLibro = False
        Exit Function
    End If
    
    ' 3. Construir la ruta completa del archivo (formato .xlsm habilitado para macros o .xlsx estándar)
    ' Usamos .xlsx por defecto para libros nuevos sin macros, o .xlsm si el libro llevará código.
    rutaCompleta = rutaCarpeta & NombreArchivo & ".xlsx"
    
    ' 4. Crear un nuevo libro de Excel (por defecto trae 1 o más hojas según tu configuración)
    Set nuevoWb = Workbooks.Add
    
    ' 5. Desactivar alertas temporalmente por si el archivo ya existe y pregunta si desea reemplazarlo
    Application.displayAlerts = False
    
    ' 6. Guardar el libro en la ruta indicada
    ' xlOpenXMLWorkbook = Formato estándar .xlsx
    nuevoWb.SaveAs Filename:=rutaCompleta, FileFormat:=xlOpenXMLWorkbook
    
    ' Reactivar alertas
    Application.displayAlerts = True
    
    ' 7. Cerrar el nuevo libro (opcional, si solo querías crearlo y guardarlo en segundo plano)
    ' nuevoWb.Close SaveChanges:=False
    
    CrearYGuardarLibro = True
    Exit Function

ErrorHandler:
    Application.displayAlerts = True
    MsgBox "Ocurrió un error al intentar guardar el archivo." & vbCrLf & _
           "Detalle: " & Err.Description, vbCritical, "Error de Guardado"
    CrearYGuardarLibro = False
End Function

Sub AbrirRecurso(cadena As String)


    On Error GoTo ManejadorError
    
    ' Validar que la cadena no esté vacía
    If Trim(cadena) = "" Then
        MsgBox "La cadena está vacía.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    MostrarProgreso
    
    ' 2. Verificar si cuenta con el prefijo de control
    If Left(cadena, 11) = "\openexcel|" Then
        ' Si es texto plano u otro tipo de comando, puedes manejarlo aquí o salir
        abrirOpenExcel cadena
        OcultarProgreso
        Exit Sub
    End If
    
    
    ' Utiliza el Explorador de Windows para abrir URLs o archivos de forma nativa
    Shell "explorer.exe """ & cadena & """", vbNormalFocus
    OcultarProgreso
    
    Exit Sub

ManejadorError:
    MsgBox "No se pudo abrir el recurso especificado." & vbCrLf & _
           "Ruta/URL: " & cadena & vbCrLf & _
           "Descripción del error: " & Err.Description, vbCritical, "Error al abrir"
           

End Sub



Sub ExportarSeleccionATexto(ByVal rutaArchivo As String)
    Dim Canal As Integer
    Dim celda As Range
    
    If rutaArchivo = "YYYYMMDD" Then
        rutaArchivo = NOMBRE_PATH_TOOL & GenerarCodigoUnico & ".txt"
    End If
    
    GuardarArchivoTexto rutaArchivo
    ' Validar que el usuario haya seleccionado un rango de celdas
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas primero.", vbExclamation, "Selección inválida"
        Exit Sub
    End If
    
    ' Obtiene un canal libre
    Canal = FreeFile
    
    MostrarProgreso
    On Error GoTo ManejadorError
    
    ' Abre el archivo en modo Output (crea o sobrescribe)
    Open rutaArchivo For Output As #Canal
    
    ' Recorre cada celda dentro de la selección actual y escribe su valor
    For Each celda In Selection
        Print #Canal, celda.Value
    Next celda
    
    ' Cierra el archivo
    Close #Canal
    
    OcultarProgreso
    
    MsgBox "¡Datos exportados con éxito!", vbInformation, "Proceso Finalizado"
    AbrirRecurso rutaArchivo
    Exit Sub

ManejadorError:
    Close #Canal
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical, "Error"
    
End Sub

Sub GuardarArchivoTexto(ByVal rutaArchivo As String)
    Dim Canal As Integer
    
    ' Obtiene un número de canal libre para manejar el archivo
    Canal = FreeFile
    
    On Error GoTo ManejadorError
    
    ' Abre (o crea) el archivo en modo escritura (Output)
    Open rutaArchivo For Output As #Canal
    
    ' Escribe el contenido en el archivo
    'Print #Canal, Contenido
    
    ' Cierra el archivo
    Close #Canal
    

    MsgBox "Archivo guardado exitosamente en: " & rutaArchivo, vbInformation, "Éxito"
    Exit Sub

ManejadorError:
    ' Asegura que el archivo se cierre si ocurre un error
    Close #Canal
    MsgBox "Ocurrió un error al guardar el archivo: " & Err.Description, vbCritical, "Error"
End Sub

Sub abrirOpenExcel(as_referencia As String)

    ' Remover el prefijo para quedarnos puramente con "libro#hoja#celda"
    Dim cuerpoRef As String
    Dim partes() As String
    Dim NombreLibro, nombreHoja, refCelda  As String
    Dim rutaCompleta As String
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    
    cuerpoRef = Mid(as_referencia, 12)
    
    ' 3. Separar usando el delimitador #
    partes = Split(cuerpoRef, "#")
    
    If UBound(partes) <> 2 Then
        MsgBox "La taxonomía está incompleta: " & as_referencia, vbCritical, "Error de Estructura"
        Exit Sub
    End If
    
    NombreLibro = partes(0)
    nombreHoja = partes(1)
    refCelda = partes(2) ' Ej: $A$5 o A5
    
    On Error GoTo ManejadorError
    
    ' 4. Gestión del Libro
    If Not LibroEstaAbierto(NombreLibro) Then
        rutaCompleta = ActiveWorkbook.Path & Application.PathSeparator & NombreLibro
        If Dir(rutaCompleta) <> "" Then
            Set wbDestino = Workbooks.Open(rutaCompleta)
        Else
            MsgBox "No se pudo encontrar el libro: " & NombreLibro, vbCritical, "Archivo no encontrado"
            Exit Sub
        End If
    Else
        Set wbDestino = Workbooks(NombreLibro)
    End If
    
    ' 5. Activación de Hoja y Celda
    wbDestino.Activate
    
    On Error Resume Next
    If Len(Trim(nombreHoja)) = 0 Then Exit Sub
    Set wsDestino = wbDestino.Worksheets(nombreHoja)
    On Error GoTo ManejadorError
    
    If wsDestino Is Nothing Then
        MsgBox "La hoja '" & nombreHoja & "' no existe en el libro " & NombreLibro, vbCritical, "Hoja no encontrada"
        Exit Sub
    End If
    
    wsDestino.Activate
    If Len(Trim(refCelda)) = 0 Then Exit Sub
    ' Limpiar símbolos de absolutez ($) si la celda los trae (ej: $A$1 -> A1) para asegurar compatibilidad con Range.Select
    refCelda = Replace(refCelda, "$", "")
    
    wsDestino.Range(refCelda).Select
    ActiveWindow.ScrollRow = wsDestino.Range(refCelda).Row
    ActiveWindow.ScrollColumn = wsDestino.Range(refCelda).Column
    
    Exit Sub

ManejadorError:
    MsgBox "Ocurrió un error al navegar: " & Err.Description, vbCritical, "Error de Navegación"
    
End Sub


