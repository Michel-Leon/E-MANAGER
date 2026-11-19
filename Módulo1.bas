Attribute VB_Name = "M�dulo1"
'-----------------------------------------------------------------
'Variables globales
'-----------------------------------------------------------------
Public Sub TPOWER():   FijarLinea "L_1": End Sub
Public Sub GFLEX_N():  FijarLinea "L_2": End Sub
Public Sub GFLEX_B():  FijarLinea "L_3": End Sub
Public Sub Blokset():  FijarLinea "L_4": End Sub

' Guarda la linea (variable + nombre oculto) y recarga los combos del Panel
Private Sub FijarLinea(ByVal linea As String)
    linea_producto = linea
    ThisWorkbook.Names.Add Name:="LineaProducto", _
                           RefersTo:="=""" & linea & """", Visible:=False
    ThisWorkbook.Worksheets("Panel").CargarCombosLinea
End Sub

' Devuelve la linea actual aunque la variable se haya borrado
Public Function LineaActual() As String
    If linea_producto = "" Then
        On Error Resume Next
        linea_producto = Replace(Mid$(ThisWorkbook.Names("LineaProducto").RefersTo, 2), """", "")
        On Error GoTo 0
    End If
    LineaActual = linea_producto
End Function

Sub abrir_formulario()

    Dim wsLogin As Worksheet
    Dim wsUsuarios As Worksheet
    Dim wsConfig As Worksheet
    Dim usuarioIngresado As String
    Dim claveIngresada As String
    Dim ultimaFila As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim nombreUsuario As String

    ' Referencias a las hojas
    Set wsLogin = ThisWorkbook.Sheets("LOGIN")
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    Set wsConfig = ThisWorkbook.Sheets("Panel")

    ' Leer usuario y clave ingresados
    usuarioIngresado = Trim(wsLogin.Range("HX83").Value)
    claveIngresada = Trim(wsLogin.Range("HX98").Value)

    ' Validar que no est?n vac?os
    If usuarioIngresado = "" Or claveIngresada = "" Then
        MsgBox "Debe ingresar usuario y contrase�a.", vbExclamation, "Login"
        Exit Sub
    End If

    ' Buscar el usuario en la columna E de USUARIOS
    encontrado = False
    ultimaFila = wsUsuarios.Cells(wsUsuarios.Rows.Count, "E").End(xlUp).Row

    For i = 1 To ultimaFila
        If Trim(wsUsuarios.Cells(i, "E").Value) = usuarioIngresado Then
            ' Usuario encontrado, ahora validar clave
            If Trim(wsUsuarios.Cells(i, "F").Value) = claveIngresada Then
                encontrado = True
                nombreUsuario = Trim(wsUsuarios.Cells(i, "D").Value)
                Exit For
            End If
        End If
    Next i

    If encontrado Then
        ' Escribir el nombre en M-Configurator, celda JO4
        wsConfig.Range("K193").Value = nombreUsuario
        'DIRIGIRSE A HOJA PANEL
        wsConfig.Activate
    Else
        MsgBox "Usuario o contrase�a incorrectos.", vbCritical, "Login"
    End If

End Sub
Sub MostrarTodasLasHojas()
    Dim ws As Worksheet
    Dim usuario As String
    Dim clave As String
    Dim intentos As Integer

    Const USUARIO_VALIDO As String = "ADMIN"
    Const CLAVE_VALIDA As String = "ADMIN2026"
    Const MAX_INTENTOS As Integer = 3

    Do
        intentos = intentos + 1

        usuario = InputBox("Ingrese el usuario:", "Acceso restringido")
        If StrPtr(usuario) = 0 Then Exit Sub   ' Cancel?

        clave = InputBox("Ingrese la contrase�a:", "Acceso restringido")
        If StrPtr(clave) = 0 Then Exit Sub     ' Cancel?

        If UCase(Trim(usuario)) = USUARIO_VALIDO And Trim(clave) = CLAVE_VALIDA Then
            Exit Do
        Else
            If intentos >= MAX_INTENTOS Then
                MsgBox "Acceso denegado. Se agotaron los intentos.", vbCritical, "Error"
                Exit Sub
            End If
            MsgBox "Usuario o contrase�a incorrectos." & vbNewLine & _
                   "Intento " & intentos & " de " & MAX_INTENTOS, vbExclamation, "Error"
        End If
    Loop

    ' Acceso concedido
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", TRUE)"
    Application.DisplayFormulaBar = True
    Application.ScreenUpdating = True

    MsgBox "Acceso concedido. Todas las hojas estan visibles.", vbInformation, "Listo"
End Sub
Sub CambiarZoomEnTodasLasHojas()
    Dim zoomStr As String
    Dim zoomVal As Integer
    Dim ws As Worksheet

    ' Solicitar el valor de zoom al usuario
    zoomStr = InputBox("Ingrese el nivel de zoom deseado (ej. 100 para 100%):", "Cambiar Zoom")
    ' Cancelar si el usuario presiona Cancelar o deja vac?o
    If zoomStr = "" Then Exit Sub

    ' Validar que el valor ingresado sea num?rico
    If Not IsNumeric(zoomStr) Then
        MsgBox "Por favor ingrese un numero valido.", vbExclamation
        Exit Sub
    End If
    zoomVal = CInt(zoomStr)
    ' Validar que el zoom est? dentro de un rango razonable
    If zoomVal < 10 Or zoomVal > 400 Then
        MsgBox "El nivel de zoom debe estar entre 10 y 400.", vbExclamation
        Exit Sub
    End If
    ' Aplicar el zoom a todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.Zoom = zoomVal
    Next ws
    MsgBox "Zoom ajustado a " & zoomVal & "% en todas las hojas.", vbInformation
End Sub
Sub OcultarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveSheet.Shapes("Mostrar Barras").Visible = True
    ActiveSheet.Shapes("ocultar Barras").Visible = False
End Sub
' Mostrar las barras y elementos de la interfaz
Sub MostarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", true)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.Shapes("ocultar Barras").Visible = True
    ActiveSheet.Shapes("Mostrar Barras").Visible = False
End Sub

