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


'=====================================================================
' MÓDULO: GENERACIÓN DE PAQUETES (FAT / APE / Cliente / Todas)
'=====================================================================

Private Const ANCHO_COLUMNA As Double = 0.38   ' unidades de columna de Excel
Private Const ALTO_FILA As Double = 3.6        ' puntos


'---------------------------------------------------------------------
' LISTAS DE HOJAS SEGÚN EL PRODUCTO (T11) Y EL TIPO DE PAQUETE
'---------------------------------------------------------------------

Private Function ObtenerHojasAPE(ByVal configuracion As String) As Variant
    Select Case configuracion
        Case "T-Power"
            ObtenerHojasAPE = Array("APE", "SIM2", "SIM3", "SIM4", "SIM5", _
                                     "PCA1", "PCA2", "PCA4", "PDD1", "ADP1", "ETQ1", "VAN1")
        Case "G-Flex"
            ObtenerHojasAPE = Array("APE", "SIM2", "SIM3", "SIM4", "SIM5", _
                                     "PCA1", "PCA2", "PCA4", "PDD1", "ADP2", "ADP4", "ETQ2", "VAN2", "ETQ4")
        Case "BlokSeT"
            ObtenerHojasAPE = Array("APE", "SIM2", "SIM3", "SIM4", "SIM5", _
                                     "PCA1", "PCA2", "PCA4", "PDD1", "ADP3", "ETQ3", "PDD2", "PCA3", "VAN3")
        Case Else
            ObtenerHojasAPE = Array()
    End Select
End Function

Private Function ObtenerHojasCliente(ByVal configuracion As String) As Variant
    Select Case configuracion
        Case "T-Power"
            ObtenerHojasCliente = Array("Cliente", "SIM2", "SIM3", "SIM4", "SIM5", _
                                         "PCA1", "PCA2", "PCA4", "PDD1", "ADP1", "ETQ1", "VAN1")
        Case "G-Flex"
            ObtenerHojasCliente = Array("Cliente", "SIM2", "SIM3", "SIM4", "SIM5", _
                                         "PCA1", "PCA2", "PCA4", "PDD1", "ADP2", "ADP4", "ETQ2", "VAN2", "ETQ4")
        Case "BlokSeT"
            ObtenerHojasCliente = Array("Cliente", "SIM2", "SIM3", "SIM4", "SIM5", _
                                         "PCA1", "PCA2", "PCA4", "PDD1", "ADP3", "ETQ3", "PDD2", "PCA3", "VAN3")
        Case Else
            ObtenerHojasCliente = Array()
    End Select
End Function

Private Function ObtenerHojasFAT() As Variant
    ObtenerHojasFAT = Array("FAT")
End Function


'---------------------------------------------------------------------
' PUNTO DE ENTRADA
' Llamar así desde el formulario, por ejemplo:
'   Case "1": GenerarPaquete 1, Me.T11.Value
'   Case "2": GenerarPaquete 2, Me.T11.Value
'   Case "3": GenerarPaquete 3, Me.T11.Value
'   Case "4": GenerarPaquete 4, Me.T11.Value
'---------------------------------------------------------------------

Public Sub GenerarPaquete(ByVal opcion As Long, ByVal configuracion As String)
    Dim hojas As Variant

    Select Case opcion
        Case 1 ' FAT
            hojas = ObtenerHojasFAT()
        Case 2 ' APE
            hojas = ObtenerHojasAPE(configuracion)
        Case 3 ' Cliente
            hojas = ObtenerHojasCliente(configuracion)
        Case 4 ' Todas
            hojas = CombinarSinDuplicados(ObtenerHojasFAT(), _
                    CombinarSinDuplicados(ObtenerHojasAPE(configuracion), ObtenerHojasCliente(configuracion)))
        Case Else
            MsgBox "Opción de paquete no reconocida.", vbExclamation
            Exit Sub
    End Select

    If Not IsArray(hojas) Then Exit Sub
    If UBound(hojas) < LBound(hojas) Then
        MsgBox "No hay hojas configuradas para '" & configuracion & "'.", vbExclamation
        Exit Sub
    End If

    CrearLibroConHojas hojas
End Sub


'---------------------------------------------------------------------
' UNE DOS ARRAYS DE NOMBRES DE HOJA SIN REPETIR (para la opción "Todas")
'---------------------------------------------------------------------

Private Function CombinarSinDuplicados(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim resultado As Collection
    Set resultado = New Collection
    Dim i As Long

    On Error Resume Next
    For i = LBound(arr1) To UBound(arr1)
        resultado.Add arr1(i), CStr(arr1(i))
    Next i
    For i = LBound(arr2) To UBound(arr2)
        resultado.Add arr2(i), CStr(arr2(i))
    Next i
    On Error GoTo 0

    Dim salida() As Variant
    ReDim salida(1 To resultado.Count)
    For i = 1 To resultado.Count
        salida(i) = resultado(i)
    Next i
    CombinarSinDuplicados = salida
End Function


'---------------------------------------------------------------------
' CREA EL LIBRO NUEVO Y COPIA CADA HOJA CONFIGURADA
'---------------------------------------------------------------------

Private Sub CrearLibroConHojas(ByVal nombresHojas As Variant)
    Dim libroOrigen As Workbook
    Set libroOrigen = ThisWorkbook

    Dim libroDestino As Workbook
    Set libroDestino = Workbooks.Add(xlWBATWorksheet)  ' libro nuevo con 1 hoja en blanco

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim i As Long
    Dim nombreHoja As String
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    Dim primeraHoja As Boolean
    primeraHoja = True

    For i = LBound(nombresHojas) To UBound(nombresHojas)
        nombreHoja = nombresHojas(i)

        Set hojaOrigen = Nothing
        On Error Resume Next
        Set hojaOrigen = libroOrigen.Worksheets(nombreHoja)
        On Error GoTo 0

        If hojaOrigen Is Nothing Then
            Debug.Print "Aviso: la hoja '" & nombreHoja & "' no existe en el libro origen y se omitió."
        Else
            If primeraHoja Then
                Set hojaDestino = libroDestino.Worksheets(1)
                hojaDestino.Name = nombreHoja
                primeraHoja = False
            Else
                Set hojaDestino = libroDestino.Worksheets.Add(After:=libroDestino.Worksheets(libroDestino.Worksheets.Count))
                hojaDestino.Name = nombreHoja
            End If

            ConfigurarYCopiarHoja hojaOrigen, hojaDestino
        End If
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Libro generado con " & libroDestino.Worksheets.Count & " hoja(s).", vbInformation
End Sub


'---------------------------------------------------------------------
' CONFIGURA ANCHO/ALTO, COPIA EL ÁREA DE IMPRESIÓN EN B2,
' DEJA SOLO VALORES SI HAY FÓRMULAS, Y AJUSTA ÁREA DE IMPRESIÓN + MÁRGENES
'---------------------------------------------------------------------

Private Sub ConfigurarYCopiarHoja(ByVal hojaOrigen As Worksheet, ByVal hojaDestino As Worksheet)
    ' 1) Ancho de columna y alto de fila para toda la hoja destino
    hojaDestino.Cells.ColumnWidth = ANCHO_COLUMNA
    hojaDestino.Cells.RowHeight = ALTO_FILA

    ' 2) Determinar el rango a copiar (área de impresión, o el usado si no hay)
    Dim rangoOrigen As Range
    If hojaOrigen.PageSetup.PrintArea = "" Then
        Set rangoOrigen = hojaOrigen.UsedRange
    Else
        Set rangoOrigen = hojaOrigen.Range(hojaOrigen.PageSetup.PrintArea)
    End If

    rangoOrigen.Copy

    Dim celdaDestino As Range
    Set celdaDestino = hojaDestino.Range("B2")

    If TieneFormulas(rangoOrigen) Then
        ' Solo valor absoluto, sin fórmulas (pero conservando el formato)
        celdaDestino.PasteSpecial Paste:=xlPasteValues
        celdaDestino.PasteSpecial Paste:=xlPasteFormats
    Else
        celdaDestino.PasteSpecial Paste:=xlPasteAll
    End If

    Application.CutCopyMode = False

    ' 3) Definir el nuevo área de impresión en el destino
    Dim nuevaArea As Range
    Set nuevaArea = hojaDestino.Range("B2").Resize(rangoOrigen.Rows.Count, rangoOrigen.Columns.Count)
    hojaDestino.PageSetup.PrintArea = nuevaArea.Address

    ' 4) Márgenes en 0
    With hojaDestino.PageSetup
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
    End With
End Sub


'---------------------------------------------------------------------
' ¿EL RANGO CONTIENE ALGUNA FÓRMULA?
'---------------------------------------------------------------------

Private Function TieneFormulas(ByVal rango As Range) As Boolean
    On Error Resume Next
    TieneFormulas = rango.HasFormula
    If Err.Number <> 0 Then TieneFormulas = False
    On Error GoTo 0
End Function