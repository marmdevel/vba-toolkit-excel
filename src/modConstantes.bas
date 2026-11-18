Attribute VB_Name = "modConstantes"
Option Explicit

' ==========================================
' CONFIGURACIÓN GENERAL Y CONSTANTES
' ==========================================

' Cambia estos valores según los nombres reales de tu archivo y tu hoja
Public Const NOMBRE_HOJA_DATOS As String = "@configuracion@"
Public Const NOMBRE_HOJA_FUNCION As String = "@funcion@"
Public Const NOMBRE_LIBRO_DATOS As String = "A-Tool.xlsm"
Public Const NOMBRE_LIBRO_PROJECTOS As String = "project.xlsx"
Public Const NOMBRE_PATH_TOOL As String = "C:\Users\PROPIETARIO\OneDrive\Documentos\Antonio\Proyectos\ToolExcel\"

' --- VARIABLES GLOBALES DEL PROYECTO ---
Public mod_ContinuaEjecucion As Boolean
Public mod_cadenaSeleccionada As String
Public SistemaInicializado As Boolean
Public mod_folderPrincipal As String

' Si el origen de los datos está en ESTE mismo libro, usa ThisWorkbook.
' Si prefieres buscarlo en otro libro abierto, puedes cambiarlo o gestionarlo dinámicamente.
Public Function ObtenerLibroDatos(as_libro) As Workbook
    ' Por defecto utiliza el libro donde corre el código (ThisWorkbook)
    Set ObtenerLibroDatos = Workbooks(as_libro)
    
    ' Si quisieras apuntar a otro libro abierto específico, podrías hacer algo como:
    ' Set ObtenerLibroDatos = Workbooks("NombreDelLibroAbierto.xlsm")
End Function
Public Function get_NOMBRE_LIBRO_DATOS()
get_NOMBRE_LIBRO_DATOS = NOMBRE_LIBRO_DATOS
End Function
Public Function get_NOMBRE_HOJA_DATOS()
get_NOMBRE_HOJA_DATOS = NOMBRE_HOJA_DATOS
End Function
Public Function get_NOMBRE_HOJA_FUNCION()
get_NOMBRE_HOJA_FUNCION = NOMBRE_HOJA_FUNCION
End Function

' --- RUTINA O FUNCIÓN DE INICIALIZACIÓN ---
Public Function InicializarSistema() As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Establecer los valores iniciales por defecto para todo el proyecto
    mod_ContinuaEjecucion = False   'variable para saber si continua ejecutando
    mod_cadenaSeleccionada = ""     'variable que devuelve la palabra seleccionada
    SistemaInicializado = True
    InicializarSistema = True
    mod_folderPrincipal = NOMBRE_PATH_TOOL
    'Workbooks(NOMBRE_LIBRO_DATOS).Path
    Exit Function

ErrorHandler:
    SistemaInicializado = False
    InicializarSistema = False
    MsgBox "Error crítico al inicializar el sistema: " & Err.Description, vbCritical, "Error de Inicio"
End Function
