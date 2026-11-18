Attribute VB_Name = "modfrmprogreso"
' Función para abrir la ventana de progreso
Sub MostrarProgreso()
    On Error Resume Next
    ' Cierra el formulario por si ya estuviera abierto de antes
    Unload frmProgreso
    On Error GoTo 0
    
    ' Muestra el formulario sin bloquear la ejecución (vbModeless)
    frmProgreso.Show vbModeless
    DoEvents
End Sub

' Función para cerrar la ventana de progreso
Sub OcultarProgreso()
    On Error Resume Next
    Unload frmProgreso
    On Error GoTo 0
End Sub
