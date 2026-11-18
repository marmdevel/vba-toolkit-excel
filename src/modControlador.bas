Attribute VB_Name = "modControlador"
Sub m_aplicarmarkdown()

    InicializarSistema
    Call CargarDatosPrevia(get_NOMBRE_LIBRO_DATOS, "@markdown@")
    
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
    'resultadoBusqueda = BuscarEnColumnas(get_NOMBRE_LIBRO_DATOS, "@funcion@", 1, 2, 3, mod_cadenaSeleccionada)
    
    ' Verificamos si arrojó resultados
    If Not IsEmpty(mod_cadenaSeleccionada) Then
        For Each c In Selection
            ls_valor = c.Value
            ls_valor = Replace(mod_cadenaSeleccionada, "activecell", ls_valor, 1, -1, vbTextCompare)
            c.Value = ls_valor
        Next
        CopiarFormatoPorValor mod_cadenaSeleccionada, get_NOMBRE_LIBRO_DATOS, "@markdown@", "C"
        'MsgBox "Encontrado:" & vbCrLf & _
        '       "Valor Columna 2 (Alias): " & resultadoBusqueda(1) & vbCrLf & _
        '       "Valor Columna 3 (Etiqueta): " & resultadoBusqueda(2), vbInformation, "Resultado de Búsqueda"
    Else
        MsgBox "No se encontró el elemento buscado.", vbExclamation, "Aviso"
    End If


End Sub





