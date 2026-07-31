Attribute VB_Name = "Enlace_libros"
Sub AbrirLineaYCopiarPortada(lineaSugerida As String)

    Dim wsConfig As Worksheet
    Dim nombreArchivo As String
    Dim rutaCompleta As String
    Dim libroDestino As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    Dim celda As Range

    Set wsConfig = ThisWorkbook.Sheets("Config_Lineas")

    ' Buscar el nombre de archivo según la línea sugerida
    Set celda = wsConfig.Columns("A").Find(lineaSugerida, LookIn:=xlValues, LookAt:=xlWhole)

    If celda Is Nothing Then
        MsgBox "No se encontró configuración para la línea: " & lineaSugerida, vbExclamation
        Exit Sub
    End If

    nombreArchivo = celda.Offset(0, 1).Value
    rutaCompleta = ThisWorkbook.Path & "\" & nombreArchivo

    If Dir(rutaCompleta) = "" Then
        MsgBox "No se encuentra el archivo: " & rutaCompleta, vbCritical
        Exit Sub
    End If

    ' Abrir el libro de la línea de producto
    Set libroDestino = Workbooks.Open(rutaCompleta)

    ' Hoja origen (en E-MANAGER)
    Set hojaOrigen = ThisWorkbook.Sheets("portada_maestra")

    ' Verificar que exista la hoja destino con el mismo nombre
    On Error Resume Next
    Set hojaDestino = libroDestino.Sheets("portada_maestra")
    On Error GoTo 0

    If hojaDestino Is Nothing Then
        MsgBox "El libro destino no tiene la hoja 'portada_maestra'.", vbExclamation
        libroDestino.Close SaveChanges:=False
        Exit Sub
    End If

    ' Limpiar contenido previo de la hoja destino (opcional)
    hojaDestino.Cells.Clear

    ' Copiar todo el rango usado y pegarlo como valores (evita romper fórmulas/formatos del destino)
    hojaOrigen.UsedRange.Copy
    hojaDestino.Range("A1").PasteSpecial xlPasteValues
    hojaDestino.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    ' Guardar el libro destino
    libroDestino.Save

    ' Traer el libro destino al frente para que el usuario lo vea
    libroDestino.Activate

    MsgBox "Datos copiados exitosamente a " & nombreArchivo, vbInformation

End Sub