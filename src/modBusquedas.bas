Attribute VB_Name = "modBusquedas"
Option Explicit

' =========================================================================
' FUNCIÓN DE BÚSQUEDA GENÉRICA MULTICOLUMNA
' =========================================================================
' Parámetros:
' - nombreLibro: Nombre del libro abierto (ej. "MiLibro.xlsm")
' - nombreHoja: Nombre de la hoja de cálculo (ej. "Datos")
' - col1, col2, col3: Pueden ser números (ej. 1, 2, 3) o letras (ej. "A", "B", "C")
' - palabraBusq: Texto exacto o parcial que se buscará en la primera columna
' Devuelve: Un Array con dos elementos (ValorCol2, ValorCol3) o Empty si no lo encuentra.
' =========================================================================
Public Function BuscarEnColumnas(ByVal NombreLibro As String, _
                                  ByVal NombreHoja As String, _
                                  ByVal col1 As Variant, _
                                  ByVal col2 As Variant, _
                                  ByVal col3 As Variant, _
                                  ByVal palabraBusq As String) As Variant
    
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error Resume Next
    Set wb = Workbooks(NombreLibro)
    If wb Is Nothing Then
        MsgBox "El libro '" & NombreLibro & "' no está abierto.", vbCritical, "Error de Libro"
        BuscarEnColumnas = Empty
        Exit Function
    End If
    
    Set ws = wb.Sheets(NombreHoja)
    If ws Is Nothing Then
        MsgBox "La hoja '" & NombreHoja & "' no existe en el libro '" & NombreLibro & "'.", vbCritical, "Error de Hoja"
        BuscarEnColumnas = Empty
        Exit Function
    End If
    On Error GoTo 0
    
    Dim lastRow As Long
    ' Determinamos la última fila con datos basándonos en la primera columna de búsqueda
    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    
    If lastRow < 1 Then
        BuscarEnColumnas = Empty
        Exit Function
    End If
    
    ' Leemos todo el rango relevante de una sola vez para máxima velocidad en memoria
    ' (Desde la fila 1 hasta la última)
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, col1), ws.Cells(lastRow, IIf(col3 > col1, col3, col1))) ' Simplificado para lectura
    
    Dim arrData As Variant
    arrData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Columns.Count)).Value ' O podemos acotar por columnas:
    
    ' Optimizamos leyendo directamente el bloque de columnas involucradas si es necesario,
    ' pero para hacerlo ultra genérico con letras/números, usamos celdas directas o un bucle seguro:
    Dim i As Long
    Dim valCol1 As String
    palabraBusq = UCase(Trim(palabraBusq))
    
    For i = 1 To lastRow
        valCol1 = UCase(CStr(ws.Cells(i, col1).Value))
        
        ' Verificamos si coincide (puedes cambiar a InStr si deseas coincidencia parcial en lugar de exacta)
        If valCol1 = palabraBusq Or (palabraBusq <> "" And InStr(valCol1, palabraBusq) > 0) Then
            
            ' Si encuentra coincidencia, extraemos los valores de la 2ª y 3ª columna indicada
            Dim resultado(1 To 2) As Variant
            resultado(1) = ws.Cells(i, col2).Value
            resultado(2) = ws.Cells(i, col3).Value
            
            BuscarEnColumnas = resultado
            Exit Function
        End If
    Next i
    
    ' Si no encuentra nada, devuelve Empty
    BuscarEnColumnas = Empty
End Function

