Attribute VB_Name = "M�dulo2"
'=====================================================================
' MÓDULO: GENERACIÓN DE PAQUETES (FAT / APE / Cliente / Todas)
'=====================================================================

Private Const ANCHO_COLUMNA As Double = 0.38        ' unidades de columna de Excel
Private Const ALTO_FILA As Double = 3.6             ' puntos
Private Const RANGO_IMPRESION As String = "B2:HA148"
Private Const RANGO_VALORES_1 As String = "E5:GX139"
Private Const RANGO_VALORES_2 As String = "E140:GR144"


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
' CONFIGURA ANCHO/ALTO, COPIA EL CONTENIDO SEGÚN EL TIPO DE HOJA,
' Y AJUSTA ÁREA DE IMPRESIÓN + CONFIGURACIÓN DE PÁGINA
'
'   - APE / FAT / Cliente: dos rangos fijos (E5:GX139 y E140:GR144)
'     se pegan SOLO COMO VALOR (sin fórmulas), porque tienen celdas
'     con fórmulas que no deben viajar al libro nuevo.
'   - Resto de hojas: se copia y pega tal cual el rango completo
'     B2:HA148 (incluye fórmulas/formato, sin cambios).
'   - En TODAS las hojas, el área de impresión final es B2:HA148,
'     con márgenes en 0, orientación horizontal y centrado en
'     ambos sentidos.
'---------------------------------------------------------------------

Private Sub ConfigurarYCopiarHoja(ByVal hojaOrigen As Worksheet, ByVal hojaDestino As Worksheet)
    ' 1) Ancho de columna y alto de fila para toda la hoja destino
    hojaDestino.Cells.ColumnWidth = ANCHO_COLUMNA
    hojaDestino.Cells.RowHeight = ALTO_FILA

    ' 2) Copiar el contenido según el tipo de hoja
    If EsHojaDePortada(hojaDestino.Name) Then
        CopiarSoloValores hojaOrigen.Range(RANGO_VALORES_1), hojaDestino.Range(RANGO_VALORES_1)
        CopiarSoloValores hojaOrigen.Range(RANGO_VALORES_2), hojaDestino.Range(RANGO_VALORES_2)
    Else
        hojaOrigen.Range(RANGO_IMPRESION).Copy
        hojaDestino.Range(RANGO_IMPRESION).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    End If

    ' 3) Área de impresión fija para todas las hojas
    hojaDestino.PageSetup.PrintArea = RANGO_IMPRESION

    ' 4) Configuración de impresión: márgenes en 0, horizontal, centrada
    With hojaDestino.PageSetup
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = True
    End With
End Sub

' ¿Es una de las 3 hojas "portada" que necesitan pegado de solo valores?
Private Function EsHojaDePortada(ByVal nombreHoja As String) As Boolean
    EsHojaDePortada = (nombreHoja = "APE" Or nombreHoja = "FAT" Or nombreHoja = "Cliente")
End Function

' Copia un rango y lo pega en el destino SOLO como valor (sin fórmulas),
' conservando el formato original.
Private Sub CopiarSoloValores(ByVal rangoOrigen As Range, ByVal rangoDestino As Range)
    rangoOrigen.Copy
    rangoDestino.PasteSpecial Paste:=xlPasteValues
    rangoDestino.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
