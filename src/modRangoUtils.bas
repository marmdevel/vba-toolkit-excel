Attribute VB_Name = "modRangoUtils"
Function f_limpiar_formato()

Selection.ClearFormats

End Function

Sub ejecutaCeldaActiva()

ls_comando = ActiveCell.Value

AbrirRecurso (CStr(ls_comando))

End Sub

Sub CopiarFormatoPorValor(valorBuscado As Variant, NombreLibro As String, NombreHoja As String, columnaBúsqueda As String)

    Dim wsOrigen As Worksheet
    Dim wbOrigen As Workbook
    Dim celdaEncontrada As Range
    
    'Dim valorBuscado As Variant
    'Dim nombreLibro As String
    'Dim nombreHoja As String
    'Dim columnaBúsqueda As String
    
    ' ==========================================
    ' CONFIGURACIÓN: Define tus parámetros aquí
    ' ==========================================
    valorBuscado = "## activecell"      ' Cambia esto por el valor que buscas
    NombreLibro = "A-Tool.xlsm"   ' Libro donde está la celda (debe estar abierto)
    NombreHoja = "@markdown@"                  ' Nombre de la hoja de origen
    columnaBúsqueda = "C"                 ' Columna donde se realizará la búsqueda
    
    ' Validar que haya una selección activa en la hoja actual
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona un rango de celdas antes de ejecutar la macro.", vbExclamation, "Advertencia"
        Exit Sub
    End If
    
    ' Validar que el libro de origen esté abierto
    On Error Resume Next
    Set wbOrigen = Workbooks(NombreLibro)
    On Error GoTo 0
    
    If wbOrigen Is Nothing Then
        MsgBox "El libro '" & NombreLibro & "' no está abierto.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Referenciar la hoja de origen
    On Error Resume Next
    Set wsOrigen = wbOrigen.Sheets(NombreHoja)
    On Error GoTo 0
    
    If wsOrigen Is Nothing Then
        MsgBox "No se encontró la hoja '" & NombreHoja & "' en el libro especificado.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Buscar el valor en la columna indicada
    Set celdaEncontrada = wsOrigen.Columns(columnaBúsqueda).Find(What:=valorBuscado, _
                                                                LookIn:=xlValues, _
                                                                LookAt:=xlWhole, _
                                                                MatchCase:=False)
    
    If Not celdaEncontrada Is Nothing Then
        ' Aplicar las propiedades de formato a la selección activa
        With Selection
            .Font.Name = celdaEncontrada.Font.Name
            .Font.Size = celdaEncontrada.Font.Size
            .Font.Bold = celdaEncontrada.Font.Bold
            .Font.Italic = celdaEncontrada.Font.Italic
            .Font.Underline = celdaEncontrada.Font.Underline
            .Font.Color = celdaEncontrada.Font.Color
            .Interior.Color = celdaEncontrada.Interior.Color
            
            ' Copiar bordes (copia los 4 bordes principales)
            'Dim i As Long
            'For i = xlEdgeLeft To xlEdgeBottom
            '    .Borders(i).LineStyle = celdaEncontrada.Borders(i).LineStyle
            '    .Borders(i).Color = celdaEncontrada.Borders(i).Color
            '    .Borders(i).Weight = celdaEncontrada.Borders(i).Weight
            'Next i
        End With
        
        MsgBox "¡Formato aplicado con éxito desde la celda " & celdaEncontrada.Address(False, False) & "!", vbInformation, "Listo"
    Else
        MsgBox "No se encontró el valor '" & valorBuscado & "' en la columna " & columnaBúsqueda & " de la hoja especificada.", vbExclamation, "No encontrado"
    End If
End Sub




Sub SeleccionarColumnasHastaUltimaFilaUsada_Global()
    Dim ws As Worksheet
    Dim UltimaFila As Long
    Dim colSet As Collection
    Dim area As Range
    Dim celda As Range
    Dim col As Long
    Dim i As Long
    
    ' Usar la hoja activa
    Set ws = ActiveSheet
    
    ' Validar que haya una celda activa
    If ActiveCell Is Nothing Then
        MsgBox "No hay celda activa.", vbExclamation
        Exit Sub
    End If
    
    ' Determinar la última fila usada en toda la hoja
    On Error Resume Next
    UltimaFila = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If UltimaFila = 0 Then
        MsgBox "La hoja está vacía.", vbInformation
        Exit Sub
    End If
    
    ' Crear colección para columnas únicas
    Set colSet = New Collection
    
    ' Recorrer todas las áreas seleccionadas (incluso no continuas)
    For Each area In Selection.Areas
        For Each celda In area.Cells
            col = celda.Column
            ' Evitar duplicados usando clave de texto
            On Error Resume Next
            colSet.Add col, CStr(col)
            On Error GoTo 0
        Next celda
    Next area
    
    ' Construir rango final con todas las columnas hasta ultimaFila global
    Dim rngFinal As Range
    For i = 1 To colSet.Count
        If rngFinal Is Nothing Then
            Set rngFinal = ws.Range(ws.Cells(1, colSet(i)), ws.Cells(UltimaFila, colSet(i)))
        Else
            Set rngFinal = Union(rngFinal, ws.Range(ws.Cells(1, colSet(i)), ws.Cells(UltimaFila, colSet(i))))
        End If
    Next i
    
    ' Seleccionar el rango final
    rngFinal.Select
End Sub

Sub DuplicarFilaOSeleccion()
    Dim rng As Range
    
    ' Verificar que lo seleccionado sea un rango válido
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecciona una celda o un rango de filas.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ManejoError
    Application.ScreenUpdating = False
    
    ' Definir el rango completo de las filas seleccionadas
    Set rng = Selection.EntireRow
    
    ' Copiar las filas seleccionadas
    rng.Copy
    
    ' Insertar las filas copiadas justo debajo de la selección, desplazando las de abajo
    rng.Offset(rng.Rows.Count, 0).EntireRow.Insert Shift:=xlDown
    
    ' Limpiar el portapapeles y restaurar la pantalla
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    MsgBox "Ocurrió un error al intentar duplicar las filas.", vbCritical
End Sub

Sub EliminarFilasVacias()
    Dim UltimaFila As Long
    Dim i As Long
    Dim ls_hoja_actual As String
    Dim ws_hoja_actual As Worksheet
    
    ' Optimizar velocidad desactivando la actualización de pantalla
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ManejadorError
    ls_hoja_actual = ActiveSheet.Name
    Set ws_hoja_actual = ActiveWorkbook.Worksheets(ls_hoja_actual)
    ' Encontrar la última fila con datos usando la celda A1 como referencia
    UltimaFila = get_ObtenerUltimaFilaSpecial(ws_hoja_actual)
    'UltimaFila = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Recorrer desde la última fila hasta la primera (de abajo hacia arriba)
    For i = UltimaFila To 1 Step -1
        ' Verificar si toda la fila está vacía
        If Application.WorksheetFunction.CountA(ActiveSheet.Rows(i)) = 0 Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i
 
ManejadorError:
    ' Reactivar configuraciones originales
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number = 0 Then
        MsgBox "¡Filas vacías eliminadas con éxito!", vbInformation, "Proceso Terminado"
    Else
        MsgBox "Ocurrió un error: " & Err.Description, vbCritical, "Error"
    End If
End Sub

Function get_ObtenerUltimaFilaSpecial(Optional Hoja As Worksheet) As Long
    Dim Fila As Long
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    On Error Resume Next
    ' Encuentra la última celda con datos/formato en la hoja
    Fila = Hoja.Cells.SpecialCells(xlCellTypeLastCell).Row
    On Error GoTo 0
    
    ' Si hay algún error o la hoja está vacía, por defecto devuelve 1
    If Fila = 0 Then Fila = 1
    
    get_ObtenerUltimaFilaSpecial = Fila
End Function
