Attribute VB_Name = "M�dulo2"
Option Explicit


' =====================================================================
'  CREAR LIBRO / ANADIR HOJAS A PARTIR DE AREAS DE IMPRESION
'  (las hojas dependen de los ComboBox: ver ReglasLibro / ReglasAdicionales)
'  Por cada area:
'    1. MAPEA las celdas con formula del area (antes de copiar).
'    2. Copia el area completa al libro destino (formatos, combinadas...).
'    3. Cada celda del mapa se reemplaza por su VALOR ABSOLUTO como TEXTO
'       (exactamente lo que se ve en la celda original).
'    4. Verifica el destino: si queda alguna formula, se vuelve a congelar.
'    5. Copia las imagenes en su misma posicion y tamano.
'    6. Configura la impresion.
'  Al final muestra un resumen por hoja: formulas convertidas / sin convertir.
' =====================================================================

' ---------------------------- CONFIGURACION ----------------------------
Private Const HOJA_CONTROL As String = "Panel"          '<-- hoja de los ComboBox
Private Const NOMBRE_RUTA As String = "RutaLibroGenerado"
Private Const ANCHO_COLUMNA As Double = 0.38
Private Const ALTO_FILA As Double = 3.6

' True  = el resultado de la formula se pega como TEXTO (igual a como se ve)
' False = se pega como valor (numero/fecha) conservando el formato
Private Const VALOR_COMO_TEXTO As Boolean = True

' Opciones de los ComboBox que no generan hoja (vacio o "No anadir")
' se ignoran solas.

' Resumen del proceso y avisos de selecciones no reconocidas
Private mInforme As String
Private mAvisos As String
Private mAreaActual As String


' =====================================================================
'  REGLAS: ComboBox -> areas de impresion
'  Formato:  "Combo :: clave => AREA1, AREA2 ; clave => AREA ; ..."
'   - clave: parte del texto de la opcion (no importan mayusculas,
'            tildes, espacios ni guiones bajos). Gana la PRIMERA que coincida.
'   - clave *: cualquier opcion que no sea "No anadir".
'   - AREA:  nombre definido o nombre de hoja con area de impresion.
'  Las hojas se crean en el orden de esta lista.
' =====================================================================
Private Function ReglasLibro() As Variant
    ReglasLibro = Array( _
        "portada      :: * => H_Portada", _
        "D_generales :: todas => HC_APE, HC_CLIENTE, HC_FAT ; APE => HC_APE ; Cliente => HC_CLIENTE ; FAT => HC_FAT", _
        "Simbologia   :: todas => SIM_2, SIM_3, SIM_4, SIM_5 ; 1 y 2 => SIM_2, SIM_3 ; 2 y 3 => SIM_3, SIM_4 ; 3 y 4 => SIM_4, SIM_5", _
        "Placa        :: * => PLG_DATOS", _
        "Cajetines    :: * => CajG_mecanico, CajG_unifilar, CajG_Control", _
        "puertas      :: TPOWER => AP_TPOWER ; GFLEX_N => AP_GFLEX_NEO ; GFLEX_B => AP_G_GFLEX ; Blokset => AP_BLOKSET", _
        "etiquetas    :: TPOWER => EQ_TPOWER ; GFLEX_N => EQ_GFLEX_NEO ; GFLEX_B => EQ_GFLEX_BASIC ; Blokset => EQ_BLOKSET", _
        "anclaje      :: Power => DA_TPOWER ; GFLEX => DA_GFLEX ; Blokset => DA_BLOKSET")
End Function

Private Function ReglasAdicionales() As Variant
    ReglasAdicionales = Array( _
        "Simbologia2  :: electrica 1 => SIM_2 ; electrica 2 => SIM_3 ; electrica 3 => SIM_4 ; electrica 4 => SIM_5", _
        "Cajetines2   :: Mecanico => CajG_mecanico ; Unifilar => CajG_unifilar ; Control => CajG_Control", _
        "Indice       :: * => Indice", _
        "D_generales2 :: APE => HC_APE ; Cliente => HC_CLIENTE ; FAT => HC_FAT")
End Function

' Lee los ComboBox y devuelve las areas a copiar, sin repetir
Private Function AreasSeleccionadas(ByVal ws As Worksheet, ByVal reglas As Variant) As Collection
    Dim res As New Collection
    Dim regla As Variant, partes As Variant, opciones As Variant, op As Variant
    Dim kv As Variant, area As Variant
    Dim combo As String, t As String, clave As String, vistos As String
    Dim encontrada As Boolean

    vistos = "|"
    For Each regla In reglas
        partes = Split(regla, "::")
        combo = Trim$(partes(0))
        t = Normalizar(TextoCombo(ws, combo))

        If t <> "" And Left$(t, 2) <> "no" Then            ' vacio o "No anadir": nada
            encontrada = False
            opciones = Split(partes(1), ";")
            For Each op In opciones
                kv = Split(op, "=>")
                clave = Normalizar(CStr(kv(0)))           ' "*" queda vacio = comodin
                If clave = "" Or InStr(1, t, clave) > 0 Then
                    For Each area In Split(kv(1), ",")
                        area = Trim$(area)
                        If InStr(vistos, "|" & Normalizar(CStr(area)) & "|") = 0 Then
                            vistos = vistos & Normalizar(CStr(area)) & "|"
                            res.Add CStr(area)
                        End If
                    Next area
                    encontrada = True
                    Exit For
                End If
            Next op
            If Not encontrada Then
                mAvisos = mAvisos & "- " & combo & ": opcion no reconocida '" & _
                          TextoCombo(ws, combo) & "'" & vbCrLf
            End If
        End If
    Next regla

    Set AreasSeleccionadas = res
End Function

' Todas las areas de las reglas (para Mapear_Formulas)
Private Function TodasLasAreas() As Collection
    Dim res As New Collection
    Dim grupo As Variant, regla As Variant, op As Variant, area As Variant
    Dim vistos As String

    vistos = "|"
    For Each grupo In Array(ReglasLibro(), ReglasAdicionales())
        For Each regla In grupo
            For Each op In Split(Split(regla, "::")(1), ";")
                For Each area In Split(Split(op, "=>")(1), ",")
                    area = Trim$(area)
                    If InStr(vistos, "|" & Normalizar(CStr(area)) & "|") = 0 Then
                        vistos = vistos & Normalizar(CStr(area)) & "|"
                        res.Add CStr(area)
                    End If
                Next area
            Next op
        Next regla
    Next grupo
    Set TodasLasAreas = res
End Function



' =====================================================================
'  MACRO 1: CREAR EL LIBRO
' =====================================================================
Public Sub CrearLibro()
    Dim wsCtrl As Worksheet, wb As Workbook
    Dim areas As Collection, n As Variant
    Dim ruta As Variant, rutaFinal As String
    Dim primera As Boolean

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)
    mAvisos = ""
    Set areas = AreasSeleccionadas(wsCtrl, ReglasLibro())

    If areas.Count = 0 Then
        MsgBox "No hay ninguna hoja seleccionada para crear el libro." & _
               IIf(mAvisos = "", "", vbCrLf & vbCrLf & mAvisos), vbExclamation
        Exit Sub
    End If

    ' Ubicacion y nombre del libro
    ruta = Application.GetSaveAsFilename( _
                InitialFileName:="Documentacion.xlsx", _
                FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
                Title:="Seleccione donde guardar el libro y escriba su nombre")
    If VarType(ruta) = vbBoolean Then Exit Sub
    rutaFinal = CStr(ruta)
    If LCase$(Right$(rutaFinal, 5)) <> ".xlsx" Then rutaFinal = rutaFinal & ".xlsx"
    If Dir(rutaFinal) <> "" Then
        If MsgBox("El archivo ya existe:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & _
                  "Desea reemplazarlo?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    On Error GoTo ErrHandler
    IniciarProceso

    Set wb = Workbooks.Add(xlWBATWorksheet)
    primera = True
    For Each n In areas
        If CopiarArea(CStr(n), wb, primera) Then primera = False
    Next n

    If primera Then
        wb.Close SaveChanges:=False
        FinalizarProceso
        MsgBox "No se copio ninguna hoja:" & vbCrLf & mInforme, vbExclamation
        Exit Sub
    End If

    RomperVinculos wb
    wb.Worksheets(1).Activate

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=rutaFinal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    GuardarRuta rutaFinal

    FinalizarProceso
    MsgBox "Libro creado:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & mInforme & _
           IIf(mAvisos = "", "", vbCrLf & "Avisos:" & vbCrLf & mAvisos), vbInformation
    Exit Sub

ErrHandler:
    FinalizarProceso
    MsgBox "Error al crear el libro en el area '" & mAreaActual & "':" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & "Completadas:" & vbCrLf & mInforme, vbCritical
End Sub


' =====================================================================
'  MACRO 2: ANADIR HOJAS ADICIONALES AL LIBRO CREADO
' =====================================================================
Public Sub AnadirHojas()
    Dim wsCtrl As Worksheet, wb As Workbook
    Dim areas As Collection, n As Variant
    Dim copiadas As Long

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)
    mAvisos = ""
    Set areas = AreasSeleccionadas(wsCtrl, ReglasAdicionales())

    If areas.Count = 0 Then
        MsgBox "No hay ninguna hoja adicional seleccionada." & _
               IIf(mAvisos = "", "", vbCrLf & vbCrLf & mAvisos), vbExclamation
        Exit Sub
    End If

    Set wb = ObtenerLibroDestino()
    If wb Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    IniciarProceso

    For Each n In areas
        If CopiarArea(CStr(n), wb, False) Then copiadas = copiadas + 1
    Next n

    If copiadas > 0 Then
        RomperVinculos wb
        wb.Save
    End If

    FinalizarProceso
    MsgBox "Hojas anadidas: " & copiadas & vbCrLf & wb.FullName & vbCrLf & vbCrLf & mInforme & _
           IIf(mAvisos = "", "", vbCrLf & "Avisos:" & vbCrLf & mAvisos), vbInformation
    Exit Sub

ErrHandler:
    FinalizarProceso
    MsgBox "Error al anadir hojas en el area '" & mAreaActual & "':" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & "Completadas:" & vbCrLf & mInforme, vbCritical
End Sub


' =====================================================================
'  MACRO 3: CREAR UN PDF CON LAS AREAS SELECCIONADAS
'  Usa las selecciones de CrearLibro y de las adicionales.
'  Arma un libro temporal con la misma copia y configuracion de
'  impresion, lo exporta a PDF y lo cierra sin guardar.
' =====================================================================
Public Sub CrearPDF()
    Dim wsCtrl As Worksheet, wb As Workbook
    Dim areas As Collection, extra As Collection, n As Variant
    Dim ruta As Variant, rutaFinal As String
    Dim primera As Boolean, vistos As String

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)
    mAvisos = ""

    ' Areas base + adicionales, sin repetir
    Set areas = AreasSeleccionadas(wsCtrl, ReglasLibro())
    Set extra = AreasSeleccionadas(wsCtrl, ReglasAdicionales())
    vistos = "|"
    For Each n In areas
        vistos = vistos & Normalizar(CStr(n)) & "|"
    Next n
    For Each n In extra
        If InStr(vistos, "|" & Normalizar(CStr(n)) & "|") = 0 Then
            vistos = vistos & Normalizar(CStr(n)) & "|"
            areas.Add CStr(n)
        End If
    Next n

    If areas.Count = 0 Then
        MsgBox "No hay ninguna hoja seleccionada para el PDF." & _
               IIf(mAvisos = "", "", vbCrLf & vbCrLf & mAvisos), vbExclamation
        Exit Sub
    End If

    ' Ubicacion y nombre del PDF
    ruta = Application.GetSaveAsFilename( _
                InitialFileName:="Documentacion.pdf", _
                FileFilter:="Archivo PDF (*.pdf), *.pdf", _
                Title:="Seleccione donde guardar el PDF y escriba su nombre")
    If VarType(ruta) = vbBoolean Then Exit Sub
    rutaFinal = CStr(ruta)
    If LCase$(Right$(rutaFinal, 4)) <> ".pdf" Then rutaFinal = rutaFinal & ".pdf"
    If Dir(rutaFinal) <> "" Then
        If MsgBox("El archivo ya existe:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & _
                  "Desea reemplazarlo?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    On Error GoTo ErrHandler
    IniciarProceso

    ' Libro temporal con las areas
    Set wb = Workbooks.Add(xlWBATWorksheet)
    primera = True
    For Each n In areas
        If CopiarArea(CStr(n), wb, primera) Then primera = False
    Next n

    If primera Then
        wb.Close SaveChanges:=False
        GuardarDatosPDF rutaFinal, areas
        FinalizarProceso
        MsgBox "No se encontro ninguna area para el PDF:" & vbCrLf & mInforme, vbExclamation
        Exit Sub
    End If

    ' Exportar todas las hojas (cada una con su area de impresion)
    mAreaActual = "exportar PDF"
    wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=rutaFinal, _
                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False

    wb.Close SaveChanges:=False
    Set wb = Nothing
    FinalizarProceso

    If MsgBox("PDF creado:" & vbCrLf & rutaFinal & vbCrLf & vbCrLf & _
              "Paginas: " & areas.Count & " area(s)" & _
              IIf(mAvisos = "", "", vbCrLf & vbCrLf & "Avisos:" & vbCrLf & mAvisos) & _
              vbCrLf & vbCrLf & "Desea abrirlo ahora?", vbYesNo + vbInformation) = vbYes Then
        ThisWorkbook.FollowHyperlink rutaFinal
    End If
    Exit Sub

ErrHandler:
    FinalizarProceso
    MsgBox "Error al crear el PDF en '" & mAreaActual & "':" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Completadas:" & vbCrLf & mInforme, vbCritical
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' =====================================================================
'  MACRO 4: ANADIR HOJAS ADICIONALES AL PDF CREADO CON CrearPDF
'  Vuelve a generar el PDF con sus areas anteriores + las adicionales
'  seleccionadas, y lo guarda encima del mismo archivo.
' =====================================================================
Public Sub AnadirAlPDF()
    Dim wsCtrl As Worksheet, wb As Workbook
    Dim rutaPDF As String, lista As String, nuevasTxt As String, vistos As String
    Dim areas As New Collection, sel As Collection
    Dim n As Variant, primera As Boolean, nNuevas As Long

    Set wsCtrl = ThisWorkbook.Worksheets(HOJA_CONTROL)

    ' PDF creado anteriormente y sus areas
    LeerDatosPDF rutaPDF, lista
    If rutaPDF = "" Or lista = "" Then
        MsgBox "No hay un PDF creado con CrearPDF." & vbCrLf & _
               "Primero cree el PDF y luego anada las hojas adicionales.", vbExclamation
        Exit Sub
    End If
    If Dir(rutaPDF) = "" Then
        MsgBox "No se encontro el PDF:" & vbCrLf & rutaPDF & vbCrLf & vbCrLf & _
               "Pudo ser movido o eliminado. Vuelva a crearlo con CrearPDF.", vbExclamation
        Exit Sub
    End If

    vistos = "|"
    For Each n In Split(lista, "|")
        If Len(n) > 0 Then
            areas.Add CStr(n)
            vistos = vistos & Normalizar(CStr(n)) & "|"
        End If
    Next n

    ' Adicionales seleccionadas que aun no estan en el PDF
    mAvisos = ""
    Set sel = AreasSeleccionadas(wsCtrl, ReglasAdicionales())
    For Each n In sel
        If InStr(vistos, "|" & Normalizar(CStr(n)) & "|") = 0 Then
            vistos = vistos & Normalizar(CStr(n)) & "|"
            areas.Add CStr(n)
            nNuevas = nNuevas + 1
            nuevasTxt = nuevasTxt & vbCrLf & "  - " & n
        End If
    Next n

    If nNuevas = 0 Then
        MsgBox "No hay hojas adicionales nuevas para anadir." & vbCrLf & _
               "(Las seleccionadas ya estan en el PDF o no hay seleccion.)" & _
               IIf(mAvisos = "", "", vbCrLf & vbCrLf & mAvisos), vbExclamation
        Exit Sub
    End If

    If MsgBox("Se anadiran al PDF:" & nuevasTxt & vbCrLf & vbCrLf & rutaPDF & vbCrLf & vbCrLf & _
              "El PDF se vuelve a generar completo con los datos actuales del libro." & _
              vbCrLf & vbCrLf & "Desea continuar?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    IniciarProceso

    Set wb = Workbooks.Add(xlWBATWorksheet)
    primera = True
    For Each n In areas
        If CopiarArea(CStr(n), wb, primera) Then primera = False
    Next n

    If primera Then
        wb.Close SaveChanges:=False
        FinalizarProceso
        MsgBox "No se encontro ninguna area para el PDF:" & vbCrLf & mInforme, vbExclamation
        Exit Sub
    End If

    mAreaActual = "exportar PDF"
    wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=rutaPDF, _
                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False
    wb.Close SaveChanges:=False
    Set wb = Nothing
    GuardarDatosPDF rutaPDF, areas
    FinalizarProceso

    If MsgBox("Hojas anadidas al PDF:" & nuevasTxt & vbCrLf & vbCrLf & rutaPDF & _
              IIf(mAvisos = "", "", vbCrLf & vbCrLf & "Avisos:" & vbCrLf & mAvisos) & _
              vbCrLf & vbCrLf & "Desea abrirlo ahora?", vbYesNo + vbInformation) = vbYes Then
        ThisWorkbook.FollowHyperlink rutaPDF
    End If
    Exit Sub

ErrHandler:
    FinalizarProceso
    MsgBox "Error al anadir al PDF en '" & mAreaActual & "':" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Si el PDF esta abierto en un visor, cierrelo e intente de nuevo.", vbCritical
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' ---- Datos del ultimo PDF (hoja oculta _DatosPDF: A1 = ruta, A2 = areas) ----
Private Function HojaDatosPDF(ByVal crear As Boolean) As Worksheet
    Dim ws As Worksheet, actual As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_DatosPDF")
    On Error GoTo 0
    If ws Is Nothing And crear Then
        Set actual = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "_DatosPDF"
        ws.Visible = xlSheetVeryHidden
        On Error Resume Next
        actual.Activate
        On Error GoTo 0
    End If
    Set HojaDatosPDF = ws
End Function

Private Sub GuardarDatosPDF(ByVal ruta As String, ByVal areas As Collection)
    Dim ws As Worksheet, n As Variant, lista As String
    For Each n In areas
        lista = lista & CStr(n) & "|"
    Next n
    Set ws = HojaDatosPDF(True)
    ws.Range("A1").Value = ruta
    ws.Range("A2").Value = lista
End Sub

Private Sub LeerDatosPDF(ByRef ruta As String, ByRef lista As String)
    Dim ws As Worksheet
    Set ws = HojaDatosPDF(False)
    If ws Is Nothing Then Exit Sub
    ruta = CStr(ws.Range("A1").Value)
    lista = CStr(ws.Range("A2").Value)
End Sub

' =====================================================================
'  ENTORNO
' =====================================================================
Private Sub IniciarProceso()
    mInforme = ""
    Application.ScreenUpdating = False
    Application.EnableEvents = False          ' evita que eventos de hoja alteren algo
    Application.Calculate                     ' valores de formulas actualizados
End Sub

Private Sub FinalizarProceso()
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' =====================================================================
'  COPIA DE UN AREA DE IMPRESION A UNA HOJA DEL LIBRO DESTINO
' =====================================================================
Private Function CopiarArea(ByVal nombreArea As String, ByVal wb As Workbook, _
                            ByVal usarPrimeraHoja As Boolean) As Boolean
    Dim rngO As Range, rngD As Range, ws As Worksheet
    Dim mapa As Range, c As Range
    Dim nMapa As Long, nOk As Long, nQuedan As Long
    Dim fallos As String

    mAreaActual = nombreArea
    Set rngO = ObtenerRango(nombreArea)
    If rngO Is Nothing Then
        mInforme = mInforme & "- " & nombreArea & ": AREA NO ENCONTRADA" & vbCrLf
        Exit Function
    End If
    Set rngO = ExpandirCombinadas(rngO)

    ' Hoja destino
    If usarPrimeraHoja Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = NombreUnico(wb, nombreArea)
    ws.Cells.ColumnWidth = ANCHO_COLUMNA
    ws.Cells.RowHeight = ALTO_FILA
    Set rngD = ws.Range("A1").Resize(rngO.Rows.Count, rngO.Columns.Count)

    wb.Activate
    ws.Activate

    ' 1) MAPA de formulas del area original
    Set mapa = MapearFormulas(rngO)
    If Not mapa Is Nothing Then nMapa = mapa.Cells.Count

    ' 2) Copia completa del area
    rngO.Copy Destination:=rngD
    Application.CutCopyMode = False
    BorrarFormas ws                            ' las imagenes se copian en el paso 5

    ' 3) Congelar cada celda del mapa
    If Not mapa Is Nothing Then
        CongelarMatricesYDesbordes mapa, rngO, rngD
        For Each c In mapa.Cells
            If CongelarCelda(c, rngO, rngD) Then nOk = nOk + 1
        Next c
    End If

    ' 4) Verificacion en el destino
    nQuedan = VerificarDestino(rngO, rngD, fallos)

    ' 5) Imagenes flotantes
    CopiarImagenes rngO, rngD

    ' 6) Impresion
    ConfigurarImpresion ws, rngD
    PrepararVista ws

    mInforme = mInforme & "- " & ws.Name & ": " & nMapa & " formulas mapeadas"
    If nQuedan = 0 Then
        mInforme = mInforme & ", todas convertidas" & vbCrLf
    Else
        mInforme = mInforme & ", " & nQuedan & " SIN convertir (" & fallos & ")" & vbCrLf
    End If
    CopiarArea = True
End Function


' Oculta la cuadricula y deja el cursor en A1 de la hoja destino.
' Nunca detiene el proceso: si Excel no deja activar la hoja, se omite.
Private Sub PrepararVista(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Parent.Activate
    ws.Activate
    If ActiveSheet.Name = ws.Name And ActiveWorkbook.Name = ws.Parent.Name Then
        ActiveWindow.DisplayGridlines = False
        Application.Goto ws.Range("A1"), True
    End If
    Err.Clear
    On Error GoTo 0
End Sub


' =====================================================================
'  MAPA DE FORMULAS
'  SpecialCells; si la hoja esta protegida (SpecialCells falla) se
'  recorre celda por celda con HasFormula.
' =====================================================================
Private Function MapearFormulas(ByVal rng As Range) As Range
    Dim r As Range, c As Range

    On Error Resume Next
    Set r = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If r Is Nothing And rng.Worksheet.ProtectContents Then
        For Each c In rng.Cells
            If c.HasFormula Then
                If r Is Nothing Then Set r = c Else Set r = Union(r, c)
            End If
        Next c
    End If

    Set MapearFormulas = r
End Function

' Formulas matriciales y desbordadas: se congelan como bloque primero,
' porque Excel no deja cambiar una sola celda de esos rangos
Private Sub CongelarMatricesYDesbordes(ByVal mapa As Range, ByVal rngO As Range, ByVal rngD As Range)
    Dim c As Range, r As Range, o As Object

    On Error Resume Next
    For Each c In mapa.Cells
        Set r = Nothing
        If c.HasArray Then
            Set r = Intersect(c.CurrentArray, rngO)
        Else
            Set o = c
            If o.HasSpill Then Set r = Intersect(o.SpillingToRange, rngO)
        End If
        If Not r Is Nothing Then
            DestinoDe(r, rngO, rngD).Value = r.Value
        End If
    Next c
    On Error GoTo 0
End Sub

' Reemplaza UNA celda del mapa por su valor en el destino.
' Devuelve True si en el destino ya no hay formula.
Private Function CongelarCelda(ByVal c As Range, ByVal rngO As Range, ByVal rngD As Range) As Boolean
    Dim src As Range, d As Range

    On Error GoTo Fallo
    If c.MergeCells Then
        ' En una combinada solo la celda principal tiene valor
        If c.Address <> c.MergeArea.Cells(1, 1).Address Then
            CongelarCelda = True
            Exit Function
        End If
        Set src = c.MergeArea
    Else
        Set src = c
    End If

    Set d = DestinoDe(src, rngO, rngD)
    If TieneImagenEnCelda(c) Then
        PegarValorConImagen src, d             ' conserva la imagen dentro de la celda
    Else
        EscribirValor c, d.Cells(1, 1)
    End If

    CongelarCelda = Not d.Cells(1, 1).HasFormula
    Exit Function
Fallo:
    CongelarCelda = False
End Function

' Escribe en "d" el valor absoluto de "c"
Private Sub EscribirValor(ByVal c As Range, ByVal d As Range)
    Dim v As Variant, t As String

    v = c.Value

    If Not VALOR_COMO_TEXTO Then
        d.Value = v
        Exit Sub
    End If

    ' Texto tal como se ve en la celda original
    t = c.Text
    If Len(t) > 0 And Replace(t, "#", "") = "" And Not IsError(v) Then
        ' La columna es angosta y Excel muestra ####: usar el formato numerico
        On Error Resume Next
        t = Application.WorksheetFunction.Text(v, c.NumberFormatLocal)
        If Err.Number <> 0 Then t = CStr(v)
        On Error GoTo 0
    End If

    If Len(t) = 0 Then
        d.Value = Empty
    Else
        ' Numeros y fechas: mantener alineacion a la derecha como en el original
        If Not IsError(v) Then
            If (IsNumeric(v) Or VarType(v) = vbDate) And VarType(v) <> vbBoolean Then
                If c.HorizontalAlignment = xlGeneral Then d.MergeArea.HorizontalAlignment = xlRight
            End If
        End If
        d.Value = "'" & t                      ' apostrofo = se guarda como texto
    End If
End Sub

' Revisa el DESTINO: cualquier formula que quede se vuelve a congelar.
' Devuelve cuantas no se pudieron convertir y sus direcciones en "fallos".
Private Function VerificarDestino(ByVal rngO As Range, ByVal rngD As Range, ByRef fallos As String) As Long
    Dim quedan As Range, c As Range, o As Range
    Dim n As Long

    On Error Resume Next
    Set quedan = rngD.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If quedan Is Nothing Then Exit Function

    For Each c In quedan.Cells
        ' Celdas ocultas dentro de una combinada: no se ven ni se imprimen
        If c.MergeCells Then
            If c.Address <> c.MergeArea.Cells(1, 1).Address Then GoTo Siguiente
        End If

        Set o = rngO.Cells(c.Row - rngD.Row + 1, c.Column - rngD.Column + 1)
        On Error Resume Next
        If TieneImagenEnCelda(o) Then
            PegarValorConImagen o.MergeArea, c.MergeArea
        Else
            EscribirValor o, c
        End If
        If c.HasFormula Then c.Value = c.Value  ' ultimo recurso
        On Error GoTo 0

        If c.HasFormula Then
            n = n + 1
            If n <= 10 Then fallos = fallos & IIf(fallos = "", "", ", ") & c.Address(0, 0)
        End If
Siguiente:
    Next c

    If n > 10 Then fallos = fallos & "..."
    VerificarDestino = n
End Function


' =====================================================================
'  AUXILIARES
' =====================================================================
Private Function TextoCombo(ByVal ws As Worksheet, ByVal nombreCombo As String) As String
    Dim v As Variant
    On Error Resume Next
    v = ws.OLEObjects(nombreCombo).Object.Text
    On Error GoTo 0
    If Not IsNull(v) Then TextoCombo = CStr(v)
End Function

Private Function TieneImagenEnCelda(ByVal c As Range) As Boolean
    Dim rico As Boolean
    On Error Resume Next
    rico = c.HasRichDataType            ' Excel 365; en versiones antiguas se ignora
    On Error GoTo 0
    TieneImagenEnCelda = rico Or (InStr(1, c.Formula, "IMAGE(", vbTextCompare) > 0)
End Function

Private Sub PegarValorConImagen(ByVal o As Range, ByVal d As Range)
    On Error Resume Next
    Err.Clear
    o.Copy
    d.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    If Err.Number <> 0 Then d.Cells(1, 1).Value = o.Cells(1, 1).Value
    On Error GoTo 0
End Sub

Private Function DestinoDe(ByVal r As Range, ByVal rngOrigen As Range, ByVal rngDestino As Range) As Range
    Set DestinoDe = rngDestino.Cells(1, 1).Offset(r.Row - rngOrigen.Row, r.Column - rngOrigen.Column) _
                              .Resize(r.Rows.Count, r.Columns.Count)
End Function

Private Function ExpandirCombinadas(ByVal rng As Range) As Range
    Dim r As Range, u As Range, c As Range, borde As Range
    Dim cambio As Boolean

    Set r = rng.Areas(1)
    Do
        cambio = False
        Set borde = Union(r.Rows(1), r.Rows(r.Rows.Count), r.Columns(1), r.Columns(r.Columns.Count))
        Set u = r
        For Each c In borde.Cells
            If c.MergeCells Then Set u = Union(u, c.MergeArea)
        Next c
        Set u = Envolvente(u)
        If u.Address <> r.Address Then
            Set r = u
            cambio = True
        End If
    Loop While cambio

    Set ExpandirCombinadas = r
End Function

Private Function Envolvente(ByVal u As Range) As Range
    Dim a As Range
    Dim f1 As Long, c1 As Long, f2 As Long, c2 As Long

    f1 = u.Areas(1).Row: c1 = u.Areas(1).Column
    f2 = f1: c2 = c1
    For Each a In u.Areas
        If a.Row < f1 Then f1 = a.Row
        If a.Column < c1 Then c1 = a.Column
        If a.Row + a.Rows.Count - 1 > f2 Then f2 = a.Row + a.Rows.Count - 1
        If a.Column + a.Columns.Count - 1 > c2 Then c2 = a.Column + a.Columns.Count - 1
    Next a
    Set Envolvente = u.Worksheet.Range(u.Worksheet.Cells(f1, c1), u.Worksheet.Cells(f2, c2))
End Function

Private Sub BorrarFormas(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub CopiarImagenes(ByVal rngOrigen As Range, ByVal rngDestino As Range)
    Dim wsO As Worksheet, wsD As Worksheet
    Dim shp As Shape, nueva As Shape
    Dim celdaO As Range, celdaD As Range

    Set wsO = rngOrigen.Worksheet
    Set wsD = rngDestino.Worksheet

    For Each shp In wsO.Shapes
        If EsImagenCopiable(shp) Then
            Set celdaO = shp.TopLeftCell
            If Not Intersect(celdaO, rngOrigen) Is Nothing Then

                ' Celda equivalente en el destino
                Set celdaD = rngDestino.Cells(1, 1).Offset(celdaO.Row - rngOrigen.Row, _
                                                          celdaO.Column - rngOrigen.Column)

                Set nueva = PegarForma(shp, wsD)
                If Not nueva Is Nothing Then
                    On Error Resume Next
                    nueva.LockAspectRatio = msoFalse
                    nueva.Left = celdaD.Left + (shp.Left - celdaO.Left)
                    nueva.Top = celdaD.Top + (shp.Top - celdaO.Top)
                    nueva.Width = shp.Width
                    nueva.Height = shp.Height
                    nueva.LockAspectRatio = shp.LockAspectRatio
                    nueva.Placement = shp.Placement
                    nueva.Name = shp.Name
                    On Error GoTo 0
                End If
            End If
        End If
    Next shp
End Sub

Private Function EsImagenCopiable(ByVal shp As Shape) As Boolean
    Select Case shp.Type
        Case msoComment, msoFormControl, msoOLEControlObject
            EsImagenCopiable = False
        Case Else
            EsImagenCopiable = shp.Visible
    End Select
End Function

Private Function DebeCopiarseComoImagen(ByVal shp As Shape) As Boolean
    Dim f As String
    If shp.Type = msoLinkedPicture Or shp.Type = msoChart Then
        DebeCopiarseComoImagen = True
        Exit Function
    End If
    On Error Resume Next
    f = shp.DrawingObject.Formula
    On Error GoTo 0
    DebeCopiarseComoImagen = (Len(f) > 0)
End Function

Private Function PegarForma(ByVal shp As Shape, ByVal wsD As Worksheet) As Shape
    Dim intento As Integer
    Dim n As Long
    Dim comoImagen As Boolean

    comoImagen = DebeCopiarseComoImagen(shp)
    n = wsD.Shapes.Count

    For intento = 1 To 5
        On Error Resume Next
        Err.Clear
        If comoImagen Then
            Application.ScreenUpdating = True        ' evita imagenes en blanco
            shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            Application.ScreenUpdating = False
        Else
            shp.Copy
        End If
        DoEvents
        wsD.Parent.Activate
        wsD.Activate
        wsD.Paste
        If Err.Number = 0 And wsD.Shapes.Count > n Then
            On Error GoTo 0
            Set PegarForma = wsD.Shapes(wsD.Shapes.Count)
            Application.CutCopyMode = False
            Exit Function
        End If
        On Error GoTo 0
        DoEvents
    Next intento

    Application.CutCopyMode = False
End Function

Private Sub ConfigurarImpresion(ByVal ws As Worksheet, ByVal rng As Range)
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = rng.Address
        .Orientation = xlLandscape
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        ' Ajustar a 1 pagina (borrar estas 3 lineas si no lo desea)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Function ObtenerRango(ByVal nombre As String) As Range
    Dim r As Range
    Dim nm As Name
    Dim sh As Worksheet
    Dim objetivo As String, nombreCorto As String

    objetivo = Normalizar(nombre)

    ' 1) Nombres definidos (se omiten los automaticos Print_Area / Print_Titles)
    For Each nm In ThisWorkbook.Names
        nombreCorto = Mid$(nm.Name, InStrRev(nm.Name, "!") + 1)
        If Normalizar(nombreCorto) = objetivo Then
            On Error Resume Next
            Set r = nm.RefersToRange
            On Error GoTo 0
            If Not r Is Nothing Then Exit For
        End If
    Next nm

    ' 2) Hoja con ese nombre -> su area de impresion
    If r Is Nothing Then
        For Each sh In ThisWorkbook.Worksheets
            If Normalizar(sh.Name) = objetivo Then
                If sh.PageSetup.PrintArea <> "" Then
                    On Error Resume Next
                    Set r = sh.Range(sh.PageSetup.PrintArea).Areas(1)
                    On Error GoTo 0
                End If
                Exit For
            End If
        Next sh
    End If

    If Not r Is Nothing Then Set r = r.Areas(1)
    Set ObtenerRango = r
End Function

Private Function Normalizar(ByVal t As String) As String
    Dim i As Long, ch As String, codigo As Long, r As String
    t = LCase$(Trim$(t))
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        codigo = AscW(ch)
        Select Case codigo
            Case 97 To 122, 48 To 57: r = r & ch           ' a-z, 0-9
            Case 224 To 229: r = r & "a"                   ' vocales con tilde
            Case 232 To 235: r = r & "e"
            Case 236 To 239: r = r & "i"
            Case 242 To 246: r = r & "o"
            Case 249 To 252: r = r & "u"
            Case 241: r = r & "n"                          ' letra enie
            ' cualquier otro caracter (espacio, _, -, ?, etc.) se ignora
        End Select
    Next i
    Normalizar = r
End Function

Private Function NombreUnico(ByVal wb As Workbook, ByVal base As String) As String
    Dim candidato As String
    Dim n As Long

    candidato = Left$(base, 31)
    n = 1
    Do While ExisteHoja(wb, candidato)
        n = n + 1
        candidato = Left$(base, 31 - Len(" (" & n & ")")) & " (" & n & ")"
    Loop
    NombreUnico = candidato
End Function

Private Function ExisteHoja(ByVal wb As Workbook, ByVal nombre As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nombre)
    On Error GoTo 0
    ExisteHoja = Not sh Is Nothing
End Function

Private Sub RomperVinculos(ByVal wb As Workbook)
    Dim links As Variant
    Dim i As Long

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            wb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If

    For i = wb.Names.Count To 1 Step -1
        If InStr(wb.Names(i).RefersTo, "[") > 0 Or InStr(wb.Names(i).RefersTo, "#REF!") > 0 Then
            wb.Names(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub GuardarRuta(ByVal ruta As String)
    ThisWorkbook.Names.Add Name:=NOMBRE_RUTA, _
        RefersTo:="=""" & Replace(ruta, """", """""") & """", Visible:=False
End Sub

Private Function LeerRuta() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(NOMBRE_RUTA).RefersTo
    On Error GoTo 0
    If Len(s) > 3 Then LeerRuta = Replace(Mid$(s, 3, Len(s) - 3), """""", """")
End Function

Private Function ObtenerLibroDestino() As Workbook
    Dim ruta As String
    Dim archivo As String
    Dim wb As Workbook
    Dim sel As Variant

    ruta = LeerRuta()
    If ruta <> "" Then
        archivo = Mid$(ruta, InStrRev(ruta, Application.PathSeparator) + 1)
        On Error Resume Next
        Set wb = Workbooks(archivo)
        On Error GoTo 0
        If wb Is Nothing Then
            If Dir(ruta) <> "" Then Set wb = Workbooks.Open(ruta)
        End If
    End If

    If wb Is Nothing Then
        sel = Application.GetOpenFilename( _
                FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
                Title:="Seleccione el libro al que desea anadir las hojas")
        If VarType(sel) = vbBoolean Then Exit Function
        Set wb = Workbooks.Open(CStr(sel))
        GuardarRuta CStr(sel)
    End If

    Set ObtenerLibroDestino = wb
End Function


' =====================================================================
'  MACRO DE APOYO: mapa de formulas, combinadas e imagenes por area
' =====================================================================
Public Sub Mapear_Formulas()
    Dim n As Variant
    Dim wbRep As Workbook, wsRep As Worksheet, ws As Worksheet
    Dim rngArea As Range, rngAmpliada As Range, celdas As Range, a As Range, c As Range
    Dim shp As Shape
    Dim fila As Long, nImg As Long, nFlot As Long
    Dim combinadas As String, cortada As String, vistos As String


    Set wbRep = Workbooks.Add(xlWBATWorksheet)
    Set wsRep = wbRep.Worksheets(1)
    wsRep.Name = "Mapa_Formulas"
    wsRep.Range("A1:J1").Value = Array("Area", "Hoja", "Rango del area", "Area ajustada", _
        "Corta combinadas", "Rango con formula", "No celdas", "Tiene combinadas", _
        "Imagenes en celda", "Imagenes flotantes")
    wsRep.Range("A1:J1").Font.Bold = True
    fila = 2
    vistos = "|"

        For Each n In TodasLasAreas()
            ' Evitar repetir un area que aparece en dos listas (ej. Oferta)
            If InStr(vistos, "|" & Normalizar(CStr(n)) & "|") = 0 Then
                vistos = vistos & Normalizar(CStr(n)) & "|"

                Set rngArea = ObtenerRango(CStr(n))
                If rngArea Is Nothing Then
                    wsRep.Cells(fila, 1).Resize(1, 3).Value = Array(CStr(n), "-", "(NO ENCONTRADA)")
                    wsRep.Cells(fila, 1).Resize(1, 10).Interior.Color = RGB(255, 199, 206)
                    fila = fila + 1
                Else
                    Set ws = rngArea.Worksheet
                    Set rngAmpliada = ExpandirCombinadas(rngArea)
                    cortada = IIf(rngAmpliada.Address = rngArea.Address, "No", "Si")

                    nFlot = 0
                    For Each shp In ws.Shapes
                        If EsImagenCopiable(shp) Then
                            If Not Intersect(shp.TopLeftCell, rngAmpliada) Is Nothing Then nFlot = nFlot + 1
                        End If
                    Next shp

                    Set celdas = Nothing
                    On Error Resume Next
                    Set celdas = rngAmpliada.SpecialCells(xlCellTypeFormulas)
                    On Error GoTo 0

                    If celdas Is Nothing Then
                        wsRep.Cells(fila, 1).Resize(1, 10).Value = Array(CStr(n), ws.Name, _
                            rngArea.Address(0, 0), rngAmpliada.Address(0, 0), cortada, _
                            "(sin formulas)", 0, "-", 0, nFlot)
                        fila = fila + 1
                    Else
                        For Each a In celdas.Areas
                            If IsNull(a.MergeCells) Then
                                combinadas = "Parcial"
                            ElseIf a.MergeCells Then
                                combinadas = "Si"
                            Else
                                combinadas = "No"
                            End If
                            nImg = 0
                            For Each c In a.Cells
                                If TieneImagenEnCelda(c) Then nImg = nImg + 1
                            Next c
                            wsRep.Cells(fila, 1).Resize(1, 10).Value = Array(CStr(n), ws.Name, _
                                rngArea.Address(0, 0), rngAmpliada.Address(0, 0), cortada, _
                                a.Address(0, 0), a.Cells.Count, combinadas, nImg, nFlot)
                            fila = fila + 1
                        Next a
                    End If
                End If
            End If
        Next n

    wsRep.Columns("A:J").AutoFit
    MsgBox "Mapa de formulas generado.", vbInformation
End Sub