Attribute VB_Name = "modDispatcher"
Option Explicit

' =========================================================================
' EJECUTOR DINÁMICO CON PARÁMETROS SEPARADOS POR ";"
' =========================================================================
Public Sub EjecutarInvocacionDinamica(ByVal resultadoBusqueda As Variant)
    
    ' 1. Validar que el arreglo no esté vacío
    If IsEmpty(resultadoBusqueda) Then
        MsgBox "No hay datos para realizar la invocación dinámica.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim nombreRutina As String
    Dim cadenaParametros As String
    
    nombreRutina = CStr(resultadoBusqueda(1))
    cadenaParametros = CStr(resultadoBusqueda(2))
    
    ' Validar que el nombre de la rutina no esté vacío
    If Trim(nombreRutina) = "" Then
        MsgBox "El nombre de la rutina/función especificado está vacío.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' 2. Dividir los parámetros usando el punto y coma (";") como separador
    Dim arrParametros() As String
    arrParametros = Split(cadenaParametros, ";")
    
    ' Limpiar espacios en blanco alrededor de cada parámetro individual
    Dim i As Long
    For i = LBound(arrParametros) To UBound(arrParametros)
        arrParametros(i) = Trim(arrParametros(i))
    Next i
    
    ' 3. Invocar dinámicamente según la cantidad de parámetros detectados usando Application.Run
    On Error GoTo ErrorHandler
    
    Dim numParams As Long
    numParams = (UBound(arrParametros) - LBound(arrParametros)) + 1
    
    ' Manejador dinámico basado en la cantidad de argumentos encontrados (soporta hasta 5 por seguridad, extensible)
    Select Case numParams
        Case 0
            Application.Run nombreRutina ' Sin parámetros
        Case 1
            If arrParametros(0) = "" Then
                Application.Run nombreRutina ' Sin parámetros
            Else
                Application.Run nombreRutina, arrParametros(0)
            End If
        Case 2
            Application.Run nombreRutina, arrParametros(0), arrParametros(1)
        Case 3
            Application.Run nombreRutina, arrParametros(0), arrParametros(1), arrParametros(2)
        Case 4
            Application.Run nombreRutina, arrParametros(0), arrParametros(1), arrParametros(2), arrParametros(3)
        Case 5
            Application.Run nombreRutina, arrParametros(0), arrParametros(1), arrParametros(2), arrParametros(3), arrParametros(4)
        Case Else
            MsgBox "La rutina '" & nombreRutina & "' tiene demasiados parámetros (máximo soportado en este bloque: 5).", vbCritical, "Límite de Parámetros"
    End Select
    
    Exit Sub

ErrorHandler:
    MsgBox "No se pudo ejecutar la rutina '" & nombreRutina & "'." & vbCrLf & _
           "Detalle del error: " & Err.Description, vbCritical, "Error de Invocación Dinámica"
End Sub

