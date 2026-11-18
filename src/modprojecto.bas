Attribute VB_Name = "modprojecto"
Const NOMBRE_HOJA_PROYECTO As String = "@project@"

Sub m_crear_proyecto()

Dim ls_libro_proyectos  As String

ls_libro_proyectos = NOMBRE_PATH_TOOL & NOMBRE_LIBRO_PROJECTOS
If LibroEstaAbierto(NOMBRE_LIBRO_PROJECTOS) Then
    Workbooks(NOMBRE_LIBRO_PROJECTOS).Activate
Else

    MsgBox "Debe abrir el libro " & NOMBRE_LIBRO_PROJECTOS
End If
End Sub
