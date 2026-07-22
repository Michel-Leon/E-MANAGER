Attribute VB_Name = "M�dulo1"
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
    Set wsConfig = ThisWorkbook.Sheets("M-Configurator")
    thisWorkbook.Sheets("Portada_Mestra").range("C3:C72").Value = ""

    ' Leer usuario y clave ingresados
    usuarioIngresado = Trim(wsLogin.Range("HX83").Value)
    claveIngresada = Trim(wsLogin.Range("HX98").Value)

    ' Validar que no estén vacíos
    If usuarioIngresado = "" Or claveIngresada = "" Then
        MsgBox "Debe ingresar usuario y contraseña.", vbExclamation, "Login"
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
        wsConfig.Range("JO4").Value = nombreUsuario

        ' Abrir el formulario
        Formulario_General.Show
    Else
        MsgBox "Usuario o contraseña incorrectos.", vbCritical, "Login"
    End If

End Sub
