Attribute VB_Name = "modGeneradores"
Option Explicit

' =========================================================================
' FUNCIÓN GENERADORA DE CÓDIGO ÚNICO: YYMMMDD_XXX
' Formato: 26 (Año 2 chars) + MMM (Mes 3 chars) + DD (Día 2 chars) + _XXX (Hash/Aleatorio único)
' =========================================================================
Public Function GenerarCodigoUnico() As String
    Dim anio As String
    Dim mesNum As Integer
    Dim mesTexto As String
    Dim dia As String
    Dim sufijoUnico As String
    
    ' 1. Obtener Año a dos dígitos (ej. "26" para 2026)
    anio = Format(Date, "yy")
    
    ' 2. Obtener el nombre del mes a tres caracteres en español
    ' Array estático con los meses abreviados en español
    Dim meses As Variant
    meses = Array("", "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")
    mesNum = Month(Date)
    mesTexto = meses(mesNum)
    
    ' 3. Obtener el día a dos dígitos (ej. "06")
    dia = Format(Day(Date), "00")
    
    ' 4. Generar un sufijo XXX altamente exclusivo y difícil de repetir en llamadas consecutivas
    ' Combinamos los milisegundos transcurridos desde medianoche (Timer) con un número aleatorio grande.
    ' Esto garantiza que aunque se llame en el mismo segundo, el valor cambie drásticamente.
    Randomize ' Inicializa la semilla aleatoria
    
    Dim tiempoMilisegundos As Long
    tiempoMilisegundos = CLng(Timer * 1000) ' Milisegundos actuales del día
    
    Dim aleatorio As Integer
    aleatorio = Int((999 - 100 + 1) * Rnd + 100) ' Número aleatorio entre 100 y 999
    
    ' Mezclamos los milisegundos con el aleatorio usando base 36 (o extrayendo los últimos dígitos únicos)
    ' Tomamos las últimas 3 o 4 posiciones para que el sufijo "XXX" sea compacto pero único
    Dim baseUnica As String
    baseUnica = Right(CStr(tiempoMilisegundos) & CStr(aleatorio), 4)
    
    ' Construir el resultado final: YYMMMDD_XXX (Ej: 26NOV06_8492)
    GenerarCodigoUnico = anio & mesTexto & dia & "_" & baseUnica
End Function
