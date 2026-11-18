Attribute VB_Name = "modBuscador"
Option Explicit

' Variable pública global donde se precargan los datos
Public allData As Variant

' 1. Precarga los datos en la variable previa antes de abrir el formulario
Sub CargarDatosPrevia(as_libro, as_hoja)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim wbDatos As Workbook
    
    ' Obtiene el libro configurado en el módulo de constantes
    Set wbDatos = ObtenerLibroDatos(as_libro)
    
    On Error Resume Next
    Set ws = wbDatos.Sheets(as_hoja)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & NOMBRE_HOJA_DATOS & "' en el libro especificado.", vbCritical, "Error"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        allData = Empty
        Exit Sub
    End If
    
    ' Carga los datos de las columnas A, C y C directamente en la variable global
    allData = ws.Range("A2:C" & lastRow).Value
End Sub

' 2. Función principal que coordina la precarga, abre el formulario y devuelve el valor
Sub MostrarBuscador()
    Dim lb_ejecutar As Boolean
    Dim ls_libro_datos As String
    lb_ejecutar = False
    ' Paso previo: Cargamos los datos en la variable global
    InicializarSistema
    
    lb_ejecutar = ExisteLibroYHoja(ActiveWorkbook.Name, get_NOMBRE_HOJA_FUNCION) And ExisteLibroYHoja(ActiveWorkbook.Name, "@configuracion@")
    If lb_ejecutar Then
        ls_libro_datos = ActiveWorkbook.Name
    Else
        ls_libro_datos = get_NOMBRE_LIBRO_DATOS
    End If
    
    Call CargarDatosPrevia(ls_libro_datos, get_NOMBRE_HOJA_DATOS)
    
    ' Validar si hay datos cargados antes de mostrar la interfaz
    If IsEmpty(allData) Then
        MsgBox "No hay datos para mostrar en el buscador.", vbExclamation, "Aviso"
        'MostrarBuscador = ""
        Exit Sub
    End If

    Dim uf As UserForm1
    Set uf = New UserForm1
    
    ' Mostrar formulario modal
    uf.Show
    
    Dim etiquetaSeleccionada As String
    On Error Resume Next
    etiquetaSeleccionada = uf.Tag
    mod_cadenaSeleccionada = etiquetaSeleccionada
    On Error GoTo 0
    
    Unload uf
    Set uf = Nothing
    ' Limpiamos la variable global al cerrar para liberar memoria
    allData = Empty
    
    Dim resultadoBusqueda As Variant
    
    If Not mod_ContinuaEjecucion Then Exit Sub
    ' Buscamos la palabra "AQUI" en el libro actual, hoja "Datos",
    ' Columna 1 (o "A") para buscar, devolviendo los valores de la Columna 2 ("B") y Columna 3 ("C")
    resultadoBusqueda = BuscarEnColumnas(ls_libro_datos, get_NOMBRE_HOJA_FUNCION, 1, 2, 3, mod_cadenaSeleccionada)
    
    ' Verificamos si arrojó resultados
    If Not IsEmpty(resultadoBusqueda) Then
        EjecutarInvocacionDinamica resultadoBusqueda
        'MsgBox "Encontrado:" & vbCrLf & _
        '       "Valor Columna 2 (Alias): " & resultadoBusqueda(1) & vbCrLf & _
        '       "Valor Columna 3 (Etiqueta): " & resultadoBusqueda(2), vbInformation, "Resultado de Búsqueda"
    Else
        MsgBox "No se encontró el elemento buscado.", vbExclamation, "Aviso"
    End If

End Sub

' 3. Lógica de filtrado multi-palabra centralizada (consumiendo la variable previa)
Sub EjecutarFiltro(ByVal textoFiltro As String, ByRef lb As MSForms.ListBox)
    If IsEmpty(allData) Then Exit Sub
    
    textoFiltro = Trim(textoFiltro)
    
    If textoFiltro = "" Then
        lb.List = allData
        Exit Sub
    End If
    
    Dim palabras As Variant
    palabras = Split(Application.WorksheetFunction.Trim(textoFiltro), " ")
    
    Dim i As Long, j As Long
    Dim filaValida As Boolean
    Dim tempArray() As Variant
    Dim matchCount As Long
    matchCount = 0
    
    ReDim tempArray(1 To UBound(allData, 1), 1 To 3)
    
    For i = 1 To UBound(allData, 1)
        Dim desc As String, aliasVal As String
        desc = UCase(CStr(allData(i, 1)))
        aliasVal = UCase(CStr(allData(i, 2)))
        
        filaValida = True
        
        For j = LBound(palabras) To UBound(palabras)
            Dim palabraBusq As String
            palabraBusq = UCase(Trim(palabras(j)))
            
            If palabraBusq <> "" Then
                If InStr(desc, palabraBusq) = 0 And InStr(aliasVal, palabraBusq) = 0 Then
                    filaValida = False
                    Exit For
                End If
            End If
        Next j
        
        If filaValida Then
            matchCount = matchCount + 1
            tempArray(matchCount, 1) = allData(i, 1)
            tempArray(matchCount, 2) = allData(i, 2)
            tempArray(matchCount, 3) = allData(i, 3)
        End If
    Next i
    
    lb.Clear
    If matchCount > 0 Then
        Dim finalArray() As Variant
        ReDim finalArray(1 To matchCount, 1 To 3)
        For i = 1 To matchCount
            finalArray(i, 1) = tempArray(i, 1)
            finalArray(i, 2) = tempArray(i, 2)
            finalArray(i, 3) = tempArray(i, 3)
        Next i
        lb.List = finalArray
    End If
    
    If lb.ListCount = 1 Then
        lb.Selected(0) = True
        'lb.SetFocus
    End If
    
End Sub


Function CargarDatosTodasLasHojas(as_workbook) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalRows As Long
    Dim currentRow As Long
    Dim tempArr As Variant
    Dim r As Long
    Dim l_allData() As Variant
    Dim wb As Workbook
    Dim esRecurso As Boolean
    Dim esPagina As Boolean
    Dim celdaA1 As String
    Dim cell As Range
    
    If Not LibroEstaAbierto(as_workbook) Then
        MsgBox "Debe abrir libro " & as_workbook
        Exit Function
    End If
    
    Set wb = Workbooks(as_workbook)
    
    ' --- PASO 1: Calcular el total EXACTO de filas no vacías ---
    totalRows = 0
    For Each ws In wb.Worksheets
        celdaA1 = Trim(CStr(ws.Range("A1").Value))
        esRecurso = (celdaA1 = "#resource" Or ws.Name = "@open@")
        esPagina = (InStr(1, celdaA1, "#pagina", vbTextCompare) > 0)
        
        If esRecurso Or esPagina Then
            'lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastRow = get_ObtenerUltimaFilaSpecial(ws)
            If lastRow >= 1 Then
                ' Contamos únicamente las filas donde la columna A tiene contenido real
                For Each cell In ws.Range("A1:A" & lastRow)
                    If Not IsEmpty(cell.Value) And cell.Value <> "" Then
                        totalRows = totalRows + 1
                    End If
                Next cell
            End If
        End If
    Next ws
    
    ' Si no hay datos en ninguna hoja válida, salimos
    If totalRows = 0 Then
        CargarDatosTodasLasHojas = Empty
        Exit Function
    End If
    
    ' Redimensionar el arreglo final con el tamaño exacto desde el principio
    ReDim l_allData(1 To totalRows, 1 To 3)
    
    ' --- PASO 2: Volcar los datos de cada hoja en el arreglo general ---
    currentRow = 1
    For Each ws In wb.Worksheets
        celdaA1 = Trim(CStr(ws.Range("A1").Value))
        esRecurso = (celdaA1 = "#resource" Or ws.Name = "@open@")
        esPagina = (InStr(1, celdaA1, "#pagina", vbTextCompare) > 0)
        
        If esRecurso Or esPagina Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow >= 1 Then
                If esPagina Then
                    ' --- CASO HOJAS #PAGINA ---
                    tempArr = ws.Range("A1:B" & lastRow).Value
                    
                    For r = LBound(tempArr, 1) To UBound(tempArr, 1)
                        If Not IsEmpty(tempArr(r, 1)) And tempArr(r, 1) <> "" Then
                            l_allData(currentRow, 1) = tempArr(r, 1) ' Columna A (Texto)
                            l_allData(currentRow, 2) = tempArr(r, 2) ' Columna B (Alias)
                            
                            ' Columna C construida automáticamente para navegación
                            'l_allData(currentRow, 3) = "\openexcel|" & wb.Name & "#" & ws.Name & "#A" & (r + 1)
                            l_allData(currentRow, 3) = "\openexcel|" & wb.Name & "#" & ws.Name & "#A" & (r)
                            
                            currentRow = currentRow + 1
                        End If
                    Next r
                    
                Else
                    ' --- CASO HOJAS #RESOURCE O @open@ ---
                    tempArr = ws.Range("A2:C" & lastRow).Value
                    
                    For r = LBound(tempArr, 1) To UBound(tempArr, 1)
                        If Not IsEmpty(tempArr(r, 1)) And tempArr(r, 1) <> "" Then
                            l_allData(currentRow, 1) = tempArr(r, 1)
                            l_allData(currentRow, 2) = tempArr(r, 2)
                            l_allData(currentRow, 3) = tempArr(r, 3)
                            currentRow = currentRow + 1
                        End If
                    Next r
                End If
            End If
        End If
    Next ws
    
    CargarDatosTodasLasHojas = l_allData
End Function

Function CargarDatosTodasLasHojas_old(as_workbook) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalRows As Long
    Dim currentRow As Long
    Dim tempArr As Variant
    Dim r As Long, c As Long
    Dim l_allData() As Variant
    
    ' --- PASO 1: Calcular el total de filas de todas las hojas ---
    totalRows = 0
    If Not LibroEstaAbierto(as_workbook) Then
        MsgBox "Debe abrir libro " & as_workbook
        Exit Function
    End If
    For Each ws In Workbooks(as_workbook).Worksheets
        ' Puedes agregar aquí un filtro si quieres excluir alguna hoja por nombre, ej: If ws.Name <> "Resumen" Then
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Si la hoja tiene datos desde la fila 2 en adelante
        If lastRow >= 2 Then
            totalRows = totalRows + (lastRow - 2 + 1) ' Cantidad de filas de A2 hasta lastRow
        End If
    Next ws
    
    ' Si no hay datos en ninguna hoja, salimos
    If totalRows = 0 Then
        CargarDatosTodasLasHojas = Empty
        Exit Function
    End If
    
    ' Redimensionar el arreglo final con el total exacto de filas y 3 columnas (A, B, C)
    ReDim l_allData(1 To totalRows, 1 To 3)
    
    ' --- PASO 2: Volcar los datos de cada hoja en el arreglo general ---
    currentRow = 1
    For Each ws In Workbooks(as_workbook).Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If ws.Range("A1").Value = "#resource" Or ws.Name = "@open@" Then
            If lastRow >= 2 Then
                ' Cargamos el rango de la hoja actual en un array temporal
                tempArr = ws.Range("A2:C" & lastRow).Value
                
                ' Copiamos cada elemento al arreglo acumulador l_allData
                For r = LBound(tempArr, 1) To UBound(tempArr, 1)
                    For c = LBound(tempArr, 2) To UBound(tempArr, 2)
                        l_allData(currentRow, c) = tempArr(r, c)
                    Next c
                    currentRow = currentRow + 1
                Next r
            End If
        End If
    Next ws
    
    ' Devolver el arreglo completo
    CargarDatosTodasLasHojas = l_allData
End Function

Sub MiMacroAsociada()
    Dim ls_libro_busqueda  As String
    Dim lb_has As Variant
    
    ls_libro_busqueda = "space.xlsx"
    ' Paso previo: Cargamos los datos en la variable global
    InicializarSistema
    lb_has = f_libro_tiene_paginas_resource(ActiveWorkbook.Name)
    If lb_has Then ls_libro_busqueda = ActiveWorkbook.Name
    allData = CargarDatosTodasLasHojas(ls_libro_busqueda)
    
    ' Validar si hay datos cargados antes de mostrar la interfaz
    If IsEmpty(allData) Then
        MsgBox "No hay datos para mostrar en el buscador.", vbExclamation, "Aviso"
        'MostrarBuscador = ""
        Exit Sub
    End If

    Dim uf As UserForm1
    Set uf = New UserForm1
    
    ' Mostrar formulario modal
    uf.Show
    
    Dim etiquetaSeleccionada As String
    On Error Resume Next
    etiquetaSeleccionada = uf.Tag
    mod_cadenaSeleccionada = etiquetaSeleccionada
    On Error GoTo 0
    
    Unload uf
    Set uf = Nothing
    ' Limpiamos la variable global al cerrar para liberar memoria
    allData = Empty
    
    Dim resultadoBusqueda As Variant
    
    If Not mod_ContinuaEjecucion Then Exit Sub
   
    ' Verificamos si arrojó resultados
    If Not IsEmpty(mod_cadenaSeleccionada) Then
        MostrarProgreso
        AbrirRecurso mod_cadenaSeleccionada
        'MsgBox "Encontrado:" & vbCrLf & _
        '       "Valor Columna 2 (Alias): " & resultadoBusqueda(1) & vbCrLf & _
        '       "Valor Columna 3 (Etiqueta): " & resultadoBusqueda(2), vbInformation, "Resultado de Búsqueda"
        OcultarProgreso
    Else
        MsgBox "No se encontró el elemento buscado.", vbExclamation, "Aviso"
    End If

End Sub

Sub mainCall()

    Dim lb_ejecutar As Boolean
    Dim ls_libro_datos As String
    lb_ejecutar = False
    ' Paso previo: Cargamos los datos en la variable global
    InicializarSistema
    
    ls_libro_datos = get_NOMBRE_LIBRO_DATOS
    
    Call CargarDatosPrevia(ls_libro_datos, get_NOMBRE_HOJA_DATOS)
    
    ' Validar si hay datos cargados antes de mostrar la interfaz
    If IsEmpty(allData) Then
        MsgBox "No hay datos para mostrar en el buscador.", vbExclamation, "Aviso"
        'MostrarBuscador = ""
        Exit Sub
    End If

    Dim uf As UserForm1
    Set uf = New UserForm1
    
    ' Mostrar formulario modal
    uf.Show
    
    Dim etiquetaSeleccionada As String
    On Error Resume Next
    etiquetaSeleccionada = uf.Tag
    mod_cadenaSeleccionada = etiquetaSeleccionada
    On Error GoTo 0
    
    Unload uf
    Set uf = Nothing
    ' Limpiamos la variable global al cerrar para liberar memoria
    allData = Empty
    
    Dim resultadoBusqueda As Variant
    
    If Not mod_ContinuaEjecucion Then Exit Sub
    ' Buscamos la palabra "AQUI" en el libro actual, hoja "Datos",
    ' Columna 1 (o "A") para buscar, devolviendo los valores de la Columna 2 ("B") y Columna 3 ("C")
    resultadoBusqueda = BuscarEnColumnas(ls_libro_datos, get_NOMBRE_HOJA_FUNCION, 1, 2, 3, mod_cadenaSeleccionada)
    
    ' Verificamos si arrojó resultados
    If Not IsEmpty(resultadoBusqueda) Then
        EjecutarInvocacionDinamica resultadoBusqueda
        'MsgBox "Encontrado:" & vbCrLf & _
        '       "Valor Columna 2 (Alias): " & resultadoBusqueda(1) & vbCrLf & _
        '       "Valor Columna 3 (Etiqueta): " & resultadoBusqueda(2), vbInformation, "Resultado de Búsqueda"
    Else
        MsgBox "No se encontró el elemento buscado.", vbExclamation, "Aviso"
    End If


End Sub

Function f_libro_tiene_paginas_resource(as_libro As String)

Dim wb_source As Workbook
Dim lb_has As Boolean
Dim ws_sheet As Worksheet
Dim ls_valor As String

Set wb_source = Workbooks(as_libro)

lb_has = False
For Each ws_sheet In wb_source.Worksheets
    ls_valor = ws_sheet.Cells(1, 1).Value
    If ls_valor = "#resource" Or (InStr(1, ls_valor, "#pagina", vbTextCompare) > 0) Then
        lb_has = True
        Exit For
    End If
Next

f_libro_tiene_paginas_resource = lb_has
End Function
