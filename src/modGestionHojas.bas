Attribute VB_Name = "modGestionHojas"
Sub DuplicarHojaActual()
    ' Duplica la hoja activa y la coloca justo después de la hoja actual
    ActiveSheet.Copy After:=ActiveSheet
End Sub

Sub AdicionarHoja()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    'ws.Name = nombreNuevaHoja
    RenombrarHojaActual
End Sub

Sub RenombrarHojaActual()
    Dim nuevoNombre As String
    Dim ws As Worksheet
    Dim nombreValido As Boolean
    Dim i As Integer
    Dim caracteresProhibidos As String
    
    ' Caracteres no permitidos por Excel en los nombres de hojas
    caracteresProhibidos = "\/?*[]:"
    
    Do
        ' Pedir el texto mediante InputBox
        nuevoNombre = InputBox("Introduce el nuevo nombre para la hoja actual:" & vbCrLf & _
                               "(Máximo 31 caracteres y sin caracteres: \ / ? * [ ] :)", "Renombrar Hoja")
        
        ' Si el usuario hace clic en Cancelar, se sale de la macro
        If StrPtr(nuevoNombre) = 0 Then Exit Sub
        
        ' Quitar espacios al inicio y al final
        nuevoNombre = Trim(nuevoNombre)
        
        ' Validación 1: Que no esté vacío
        If nuevoNombre = "" Then
            MsgBox "El nombre no puede estar vacío.", vbCritical, "Error"
            nombreValido = False
        
        ' Validación 2: Longitud máxima de 31 caracteres
        ElseIf Len(nuevoNombre) > 31 Then
            MsgBox "El nombre es demasiado largo. Excel permite un máximo de 31 caracteres.", vbCritical, "Error de sintaxis"
            nombreValido = False
            
        Else
            ' Validación 3: Revisar caracteres prohibidos
            nombreValido = True
            For i = 1 To Len(caracteresProhibidos)
                If InStr(1, nuevoNombre, Mid(caracteresProhibidos, i, 1)) > 0 Then
                    MsgBox "El nombre contiene caracteres no permitidos." & vbCrLf & _
                           "Evita usar: \ / ? * [ ] :", vbCritical, "Error de sintaxis"
                    nombreValido = False
                    Exit For
                End If
            Next i
            
            ' Validación 4: Verificar que no exista otra hoja con el mismo nombre
            If nombreValido Then
                For Each ws In ActiveWorkbook.Sheets
                    ' Compara sin distinguir mayúsculas/minúsculas y excluye la hoja actual
                    If ws.Name <> ActiveSheet.Name And LCase(ws.Name) = LCase(nuevoNombre) Then
                        MsgBox "Ya existe otra hoja con ese nombre en este libro. Elige uno diferente.", vbCritical, "Nombre duplicado"
                        nombreValido = False
                        Exit For
                    End If
                Next ws
            End If
        End If
        
    ' El bucle se repite hasta que el usuario introduzca un nombre válido
    Loop Until nombreValido = True
    
    ' Aplicar el nuevo nombre a la hoja activa
    On Error GoTo ManejadorError
    ActiveSheet.Name = nuevoNombre
    Exit Sub

ManejadorError:
    MsgBox "Ocurrió un error inesperado al renombrar la hoja: " & Err.Description, vbCritical, "Error"
End Sub

Function ExisteLibroYHoja(ByVal NombreLibro As String, ByVal nombreHoja As String) As Boolean
    Dim wkb As Workbook
    Dim libroExiste As Boolean
    Dim hojaExiste As Boolean
    
    ExisteLibroYHoja = False
    libroExiste = False
    hojaExiste = False
    
    ' 1. Verificar si el libro está abierto en la sesión de Excel
    On Error Resume Next
    Set wkb = Workbooks(NombreLibro)
    On Error GoTo 0
    
    If wkb Is Nothing Then Exit Function ' Si el libro no está abierto, retorna False
    
    ' 2. Verificar si la hoja existe dentro de ese libro
    On Error Resume Next
    hojaExiste = Not (wkb.Sheets(nombreHoja) Is Nothing)
    On Error GoTo 0
    
    ' Si ambos existen, la función devuelve True
    If hojaExiste Then
        ExisteLibroYHoja = True
    End If
End Function

Function LibroEstaAbierto(ByVal NombreArchivo As String) As Boolean
    Dim wb As Workbook
    Dim Abierto As Boolean
    
    Abierto = False
    
    ' Recorre todos los libros abiertos en la sesión actual de Excel
    For Each wb In Workbooks
        ' Compara el nombre del archivo ignorando mayúsculas/minúsculas
        If LCase(wb.Name) = LCase(NombreArchivo) Then
            Abierto = True
            Exit For
        End If
    Next wb
    
    LibroEstaAbierto = Abierto
End Function

Sub m_abrir_notebook()

Dim ls_nombre_notebook As String

ls_nombre_notebook = "notebook.xlsx"
If LibroEstaAbierto(ls_nombre_notebook) Then
    Workbooks(ls_nombre_notebook).Activate
Else
    MsgBox "Debe abrir libro " & ls_nombre_notebook
End If

End Sub
