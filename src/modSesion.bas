Attribute VB_Name = "modSesion"

Sub RegistrarLibrosAbiertos()
    Dim wbAbierto As Workbook
    Dim wsDestino As Worksheet
    Dim UltimaFila As Long
    Dim filaDestino As Long
    Dim fechaHoy As String
    Dim NombreLibro As String
    Dim rutaLibro As String
    Dim yaRegistrado As Boolean
    Dim i As Long
    
    ' -------------------------------------------------------------------------
    ' CONFIGURACIÓN: Cambia "Resumen" por el nombre de tu hoja de destino
    ' y "EsteLibro" por el nombre del archivo donde corre la macro (si aplica).
    ' -------------------------------------------------------------------------
    On Error Resume Next
    
    Set wsDestino = Workbooks("space.xlsx").Sheets("@open@") ' <--- Cambia "Resumen" por tu hoja
    On Error GoTo 0
    
    If wsDestino Is Nothing Then
        MsgBox "No se encontró la hoja de destino especificada.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Formatear la fecha actual como texto 'yyyy-mm-dd'
    fechaHoy = Format(Date, "yyyy-mm-dd")
    
    ' Asegurar que la columna B tenga formato de texto antes de escribir
    wsDestino.Columns("B").NumberFormat = "@"
    
    ' Encontrar la última fila con datos en la hoja de destino
    UltimaFila = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
    If UltimaFila = 0 Then UltimaFila = 1
    
    filaDestino = UltimaFila + 1
    'If wsDestino.Cells(1, 1).Value = "" Then filaDestino = 1 ' Si la hoja está vacía
    
    ' Recorrer todos los libros abiertos en la sesión de Excel
    For Each wbAbierto In Application.Workbooks
        ' Evitar registrar el propio libro que ejecuta la macro (opcional, puedes quitarlo si deseas)
        If wbAbierto.Name <> ActiveWorkbook.Name Then
            
            NombreLibro = wbAbierto.Name
            
            ' Validar si el libro ha sido guardado en disco (tiene ruta) o es nuevo sin guardar
            On Error Resume Next
            rutaLibro = wbAbierto.FullName
            On Error GoTo 0
            
            If rutaLibro = "" Or Err.Number <> 0 Then
                rutaLibro = "Libro no guardado en disco"
            End If
            
            ' Verificar si ya existe un registro para este libro en la fecha de hoy
            yaRegistrado = False
            For i = 1 To UltimaFila
                If wsDestino.Cells(i, 1).Value = NombreLibro And wsDestino.Cells(i, 2).Value = fechaHoy Then
                    yaRegistrado = True
                    Exit For
                End If
            Next i
            
            ' Si NO está registrado para hoy, lo agregamos
            If Not yaRegistrado Then
                ' Si la celda de la columna B está vacía en la fila destino, aseguramos texto
                wsDestino.Cells(filaDestino, 1).Value = NombreLibro
                wsDestino.Cells(filaDestino, 2).NumberFormat = "@"
                wsDestino.Cells(filaDestino, 2).Value = fechaHoy
                wsDestino.Cells(filaDestino, 3).Value = "ms-excel:ofe|u|" & rutaLibro
                
                filaDestino = filaDestino + 1
                UltimaFila = ultimaFil + 1 ' Actualizamos para el bucle de validación interno si hubiera duplicados en la misma ejecución
            End If
            
        End If
    Next wbAbierto
    
    MsgBox "Proceso completado con éxito.", vbInformation, "Registro de Libros"
End Sub


