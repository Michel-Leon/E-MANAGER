VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formulario_General 
   Caption         =   "Datos de entrada"
   ClientHeight    =   10790
   ClientLeft      =   50
   ClientTop       =   -150
   ClientWidth     =   12410
   OleObjectBlob   =   "Formulario_General.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formulario_General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

private sub UserForm_Initialize()
   'Tamaño de la ventana
      Me.Height = 725
      Me.Width = 1005  
   ' Configuracion listas desplegables
      with Me.T01
         .AddItem "Contrato"
         .AddItem "Polizas"
         .AddItem "Orden de compra"
         .AddItem "No Aplica"
      end with
      With Me.T02
         .AddItem "Suministro"
         .AddItem "Instalación"
         .AddItem "Acompañamiento"
         .AddItem "Pruebas SAT"
         .AddItem "No Aplica"
      end with
      with Me.T05
         .AddItem "Distribucion"
         .AddItem "Trasnferencia"
         .AddItem "Compensacion"
         .AddItem "No Aplica"
      end with
      with Me.T08
         .AddItem "Presencial"
         .AddItem "Virtual"
         .AddItem "Registro Fotografico"
         .AddItem "Video Operativo"
         .AddItem "Parametrización"
      end with
      '==================================================
      with Me.T11
         .AddItem "T-Power"
         .AddItem "G-Flex"
         .AddItem "BlokSeT"
         .AddItem "D-Critical"
         .AddItem "S-Energy"
         .AddItem "M-Board"
         .AddItem "G-Sync"
         .AddItem "A-Tranfer"
      end with
      With Me.T16
         .AddItem "RAL 7035"
         .AddItem "RAL 7032"
         .AddItem "RAL 7011"
         .AddItem "RAL 7016"
         .AddItem "RAL 9001"
         .AddItem "RAL 9003"
         .AddItem "RAL 7035 Liso"
         .AddItem "RAL 7032 Liso"
         .AddItem "RAL 9001 Liso"
         .AddItem "RAL 9003 Liso"
         .AddItem "Azul"
         .AddItem "Negro Texturizado"
         .AddItem "Verde maquina"
         .AddItem "Pendiente"
         .AddItem "No Aplica"
      end with
      with Me.T17
         .additem "Galvanizado"
         .additem "Inoxidable"
      end with
      with Me.T18 
         .AddItem "14" 
         .AddItem "16"
      end with
      with Me.T19 
         .AddItem "20"
         .AddItem "30"
         .AddItem "31"
         .AddItem "42"
         .AddItem "54"
         .AddItem "55"
         .AddItem "65"
         .AddItem "66"
         .AddItem "No Aplica"
      end with
      With Me.T111
         .AddItem "Forma 1"
         .AddItem "Forma 2a"
         .AddItem "Forma 2b"
         .AddItem "Forma 3a"
         .AddItem "Forma 3b"
         .AddItem "Forma 4a"
         .AddItem "Forma 4b"
         .AddItem "No Aplica"
      end with
      with Me.T112
         .AddItem "Zocalo"
         .AddItem "Base estructural"
         .AddItem "No Aplica"
      end with
      with Me.T113
         .AddItem "Armella"
         .AddItem "Angulo de izaje"
         .AddItem "No Aplica"
      end with
      with Me.T114
         .AddItem "Poste"
         .AddItem "Pared"
         .AddItem "No Aplica"
      end with
      '==================================================
      with Me.T22
         .AddItem "10"
         .AddItem "15"
         .AddItem "20"
         .AddItem "25"
         .AddItem "30"
         .AddItem "35"
         .AddItem "45"
         .AddItem "55"
         .AddItem "65"
         .AddItem "85"
         .AddItem "100"
         .AddItem "No Aplica"
      end with
      With Me.T23
         .AddItem "110 VAC"
         .AddItem "120 VAC"
         .AddItem "208 VAC"
         .AddItem "220 VAC"
         .AddItem "380 VAC"
         .AddItem "400 VAC"
         .AddItem "440 VAC"
         .AddItem "460 VAC"
         .AddItem "480 VAC"
         .AddItem "690 VAC"
         .AddItem "800 VAC"
         .AddItem "1000 VAC"
         .AddItem "308/208 VAC"
         .AddItem "440/208 VAC"
         .AddItem "440/220 VAC"
         .AddItem "460/208 VAC"
         .AddItem "460/220 VAC"
         .AddItem "480/220 VAC"
         .AddItem "480/208 VAC"
         .AddItem "220 VAC/125 VDC"
         .AddItem "440 VAC/125 VDC"
         .AddItem "No Aplica"
      end with
      with Me.T24
         .AddItem "3F+N+T"
         .AddItem "3F+N"
         .AddItem "3F+T"
         .AddItem "2F+N+T"
         .AddItem "2F+N"
         .AddItem "2F+T"
         .AddItem "1F+N"
         .AddItem "1F+N+T"
         .AddItem "No Aplica"
      end with
      with Me.T25
         .AddItem "ENEL CODENSA"
         .AddItem "EPM"
         .AddItem "ESSA"
         .AddItem "No Aplica"
      end with
      with Me.T28
         .AddItem "Superior"
         .AddItem "Inferior"
         .AddItem "Superior / Inferior"
      end with
      with Me.T29
         .AddItem "Superior"
         .AddItem "No Aplica"
      end with
      with Me.T210
         .AddItem "85"
      end with
      '==================================================
      With Me.T31
         .AddItem "Barril"
         .AddItem "Socomec 2P"
         .AddItem "Socomec 3P"
         .AddItem "Escalera"
         .AddItem "No Aplica"
      end with
      with Me.T32
         .AddItem "Desnuda"
         .AddItem "Estaño"
         .AddItem "Recubrimiento en plata"
         .AddItem "Funda Termoencogible"
         .AddItem "Estaño + Funda"
         .AddItem "Recubrimiento en plata + Funda"
         .AddItem "No Aplica"
      end with
      with Me.T33
         .AddItem "50%"
         .AddItem "100%"
      end with
      with Me.T35
         .AddItem "Soldador"
         .AddItem "SIS"
         .AddItem "TWK"
         .AddItem "XHHW"
         .AddItem "No Aplica"
      end with
      with Me.T36
         .AddItem "Aplica"
         .AddItem "No aplica"
      end with
      with Me.T37
         .AddItem "No Aplica"
         .AddItem "1"
         .AddItem "2"
         .AddItem "3"
         .AddItem "4"
         .AddItem "5"
         .AddItem "6"
         .AddItem "7"
         .AddItem "8"
         .AddItem "9"
         .AddItem "10"
         .AddItem "Mas de 10"
      end with
      '==================================================
      with Me.T41
         .AddItem "12"
         .AddItem "24"
         .AddItem "120"
         .AddItem "125"
         .AddItem "208"
         .AddItem "220"
         .AddItem "440"
         .AddItem "No Aplica"
      end with
      with Me.T42
         .AddItem "12"
         .AddItem "24"
         .AddItem "120"
         .AddItem "125"
         .AddItem "208"
         .AddItem "220"
         .AddItem "440"
         .AddItem "No Aplica"
      end with
      With Me.T43
         .AddItem "Sector Petrolero"
         .AddItem "Sector Industrial"
      end with
      With Me.T44
         .AddItem "SIS"
         .AddItem "TWK"
         .AddItem "XHHW"
         .AddItem "No Aplica"
      end with
      with Me.T45
         .AddItem "Alfa Numerico"
         .AddItem "Origen / Destino"
      end with
      with Me.T46
         .AddItem "Negro-Blanco-Verde"
         .AddItem "Rojo-Azul-Verde"
         .AddItem "Rojo-Negro-Verde"
         .AddItem "Rojo-Blanco-Verde"
         .AddItem "No Aplica"
      end with
      with Me.T47
         .AddItem "Aplica"
         .AddItem "No Aplica"
      end with
      with Me.T48
         .AddItem "Alfa Numerico"
         .AddItem "Origen / Destino"
         .AddItem "No Aplica"
      end with
      '==================================================
      with Me.T51
         .AddItem "Aplica"
         .AddItem "No Aplica"
      end with
      '==================================================
      with Me.T61
         .AddItem "Tipo I"
         .AddItem "Tipo II"
         .AddItem "Tipo III"
         .AddItem "Tipo IV"
         .AddItem "Tipo V"
      end with
      '==================================================
      with Me.T71
         .AddItem "Rutina"
         .AddItem "Especial"
         .AddItem "Avanzada"
         .AddItem "Premium"
      end with
      '==================================================
      with Me.T81
         .AddItem "Aplica"
         .AddItem "No Aplica"
      end with
      With Me.T82
         .AddItem "Gravoply"
         .AddItem "Adeshivo"
         .AddItem "Metálica"
      end with
      With Me.T83
         .AddItem "Gravoply"
         .AddItem "Adhesivo gris"
         .AddItem "Metálica"
         .AddItem "No aplica"
      end with
      With Me.T84
         .AddItem "Gravoply"
         .AddItem "Gris"
      end with
      With Me.T85
         .AddItem "Gravoply"
         .AddItem "Vinilo"
      end with
      With Me.T86
         .AddItem "Gris"
         .AddItem "Blanco"
         .AddItem "Negro"
      end with
      With Me.T87
         .AddItem "Aplica"
         .AddItem "No Aplica"
      end with
      With Me.T88
         .AddItem "Gris"
         .AddItem "Blanco"
         .AddItem "Negro"
         .AddItem "No Aplica"
      end with
      With Me.T89
         .AddItem "Aplica"
         .AddItem "No Aplica"
      end with
      With Me.T810
         .AddItem "Gravoply"
         .AddItem "Galvanizado"
         .AddItem "Inoxidable"
         .AddItem "No Aplica"
      end with
end sub    
Private Sub T11_Change()
   Dim configuracion As String
   configuracion = Me.T11.Value
   
   With Me.T12
      .Clear
      Select Case configuracion
         Case "T-Power"
            .AddItem "1600"
            .AddItem "1800"
            .AddItem "2000"
            .AddItem "2200"
         Case "G-Flex"
            .AddItem "600"
            .AddItem "700"
            .AddItem "800"
            .AddItem "900"
            .AddItem "1100"
         Case "BlokSeT"
            .AddItem "por definir"
         Case "D-Critical"
            .AddItem "por definir"
         Case "S-Energy"
            .AddItem "por definir"
         Case "M-Board"
            .AddItem "por definir"
         Case "G-Sync"
            .AddItem "por definir"
         Case "A-Tranfer"
            .AddItem "por definir"
      End Select
   End With

   With Me.T15
      .Clear
      Select Case configuracion
         Case "T-Power"
            .AddItem "400"
            .AddItem "600"
            .AddItem "800"
            .AddItem "1000"
         Case "G-Flex"
            .AddItem "300"
            .AddItem "400"
         Case "BlokSeT"
            .AddItem "por definir"
         Case "D-Critical"
            .AddItem "por definir"
         Case "S-Energy"
            .AddItem "por definir"
         Case "M-Board"
            .AddItem "por definir"
         Case "G-Sync"
            .AddItem "por definir"
         Case "A-Tranfer"
            .AddItem "por definir"
      End Select
   End With

   CalcularEspacios
   
   Select Case configuracion
      Case "T-Power", "D-Critical", "S-Energy", "M-Board", "G-Sync", "A-Tranfer"
         Me.T114.Enabled = False
         Me.T210.Enabled = False
         Me.T29.Enabled = True
      Case "G-Flex"
         Me.T114.Enabled = True
         Me.T210.Enabled = False
         Me.T29.Enabled = False
      Case "BlokSeT"
         Me.T114.Enabled = False
         Me.T210.Enabled = True
         Me.T29.Enabled = False
   End Select
End Sub

Private Sub T13_Change()
   CalcularEspacios
End Sub
Private sub T23_Change()
   DIM configuracion as string
   configuracion = Me.T23.Value
   Select Case configuracion
      Case "110 VAC"
         Me.L1.Caption = "Negro"
         Me.L2.Caption = "-"
         Me.L3.Caption = "-"
         Me.L4.Caption = "Blanco"
         Me.L5.Caption = "Verde"
      Case "120 VAC"
         Me.L1.Caption = "Amarillo"
         Me.L2.Caption = "-"
         Me.L3.Caption = "-"
         Me.L4.Caption = "Blanco"
         Me.L5.Caption = "Verde"
      Case "208 VAC","220 VAC"
         Me.L1.Caption = "Amarillo"
         Me.L2.Caption = "Azul"
         Me.L3.Caption = "Rojo"
         Me.L4.Caption = "Blanco"
         Me.L5.Caption = "Verde"
      Case "380 VAC"
         Me.L1.Caption = "Cafe"
         Me.L2.Caption = "Negro"
         Me.L3.Caption = "Amarillo"
         Me.L4.Caption = "Blanco"
         Me.L5.Caption = "Verde"
      Case "400 VAC","440 VAC","460 VAC","480 VAC","690 VAC","800 VAC","1000 VAC","308/208 VAC","220 VAC/125 VDC","440 VAC/125 VDC"
         Me.L1.Caption = "Cafe"
         Me.L2.Caption = "Naranja"
         Me.L3.Caption = "Amarillo"
         Me.L4.Caption = "Gris"
         Me.L5.Caption = "Verde"
      Case "440/208 VAC","440/220 VAC","460/208 VAC","460/220 VAC","480/220 VAC","480/208 VAC"
         Me.L1.Caption = "Café/Amarillo" 
         Me.L2.Caption = "Naranja/Azul"
         Me.L3.Caption = "Amarillo/Rojo"
         Me.L4.Caption = "Gris/Blanco"
         Me.L5.Caption = "Verde"  
      case else
         Me.L1.Caption = "-"
         Me.L2.Caption = "-"
         Me.L3.Caption = "-"
         Me.L4.Caption = "-"
         Me.L5.Caption = "-"
   End Select
End Sub
Sub CalcularEspacios()

   Dim valor As Double
   Dim resultado As String
   Dim medidas() As Variant

   Select Case Trim(T11.Value)

      Case "G-Flex"
         ' --- G-Flex: solo UN ancho a la vez, no es suma de combinaciones ---
         medidas = Array(1120, 720, 620, 520, 400)

         If Not IsNumeric(T13.Value) Then
               T14.Value = ""
               Exit Sub
         End If

         valor = Val(T13.Value)

         If valor <= 0 Then
               T14.Value = ""
               Exit Sub
         End If

         Dim i As Long
         Dim esValido As Boolean
         esValido = False

         For i = LBound(medidas) To UBound(medidas)
               If valor = medidas(i) Then
                  esValido = True
                  Exit For
               End If
         Next i

         If esValido Then
               ' Ancho correcto, no se necesita mensaje
               T14.Value = ""
         Else
               ' No coincide con ninguno: mostrar cuáles SÍ son válidos + el más cercano
               Dim mejor As Variant
               Dim menorDif As Double
               Dim dif As Double

               menorDif = -1
               For i = LBound(medidas) To UBound(medidas)
                  dif = Abs(valor - medidas(i))
                  If menorDif = -1 Or dif < menorDif Then
                     menorDif = dif
                     mejor = medidas(i)
                  End If
               Next i

               T14.Value = "Ancho no válido para G-Flex." & vbNewLine & _
                  "Anchos permitidos: 400 -- 520 -- 620 -- 720 -- 1120" & vbNewLine & _
                  "El más cercano a " & valor & " es: " & mejor
         End If

         Exit Sub

      Case "T-Power"
         medidas = Array(1120, 720, 400, 200) ' orden descendente

      Case Else
         T14.Value = "Seleccione un tipo válido (T-Power o G-Flex)"
         Exit Sub
   End Select

   ' --- A partir de aquí solo llega T-Power (lógica de combinaciones sin cambios) ---

   If Not IsNumeric(T13.Value) Then
      T14.Value = ""
      Exit Sub
   End If

   valor = Val(T13.Value)

   If valor <= 0 Then
      T14.Value = ""
      Exit Sub
   End If

   Dim combinaciones As Collection
   Set combinaciones = New Collection

   BuscarCombinaciones medidas, 0, valor, "", combinaciones

   Dim j As Long
   resultado = ""
   For j = 1 To combinaciones.Count
      If resultado <> "" Then resultado = resultado & vbNewLine
      resultado = resultado & combinaciones(j)
   Next j

   If resultado = "" Then
      T14.Value = "No hay combinación exacta posible"
   Else
      T14.Value = resultado
   End If

End Sub

Private Sub BuscarCombinaciones(medidas() As Variant, indiceInicio As Long, restante As Double, actual As String, combinaciones As Collection)

    Dim i As Long
    Dim nuevoActual As String

    If restante = 0 Then
        If actual <> "" Then
            If EsCombinacionValida(actual) Then
                combinaciones.Add OrdenarAscendente(actual)
            End If
        End If
        Exit Sub
    End If

    If restante < 0 Then Exit Sub

    For i = indiceInicio To UBound(medidas)
        If actual = "" Then
            nuevoActual = CStr(medidas(i))
        Else
            nuevoActual = actual & "+" & medidas(i)
        End If
        BuscarCombinaciones medidas, i, restante - medidas(i), nuevoActual, combinaciones
    Next i

End Sub

' Valida que:
' 1) Solo pueda haber máximo un 200 en la combinación
' 2) Si hay un 200, debe estar acompañado de al menos un 400 o un 720
Private Function EsCombinacionValida(cadena As String) As Boolean

    Dim partes() As String
    Dim i As Long 
    Dim conteo200 As Long
    Dim tieneSoporte As Boolean ' 400 o 720

    partes = Split(cadena, "+")
    conteo200 = 0
    tieneSoporte = False

    For i = LBound(partes) To UBound(partes)
        Select Case CLng(partes(i))
            Case 200
                conteo200 = conteo200 + 1
            Case 400, 720
                tieneSoporte = True
        End Select
    Next i

    ' Más de un 200 en la combinación -> inválida
    If conteo200 > 1 Then
        EsCombinacionValida = False
        Exit Function
    End If

    ' Un 200 sin 400 ni 720 de soporte -> inválida
    If conteo200 = 1 And Not tieneSoporte Then
        EsCombinacionValida = False
        Exit Function
    End If

    EsCombinacionValida = True

End Function

Private Function OrdenarAscendente(cadena As String) As String
   Dim partes() As String
   Dim numeros() As Long
   Dim i As Long, j As Long, temp As Long

   partes = Split(cadena, "+")
   ReDim numeros(LBound(partes) To UBound(partes))

   For i = LBound(partes) To UBound(partes)
      numeros(i) = CLng(partes(i))
   Next i

   For i = LBound(numeros) To UBound(numeros) - 1
      For j = i + 1 To UBound(numeros)
         If numeros(i) > numeros(j) Then
               temp = numeros(i)
               numeros(i) = numeros(j)
               numeros(j) = temp
         End If
      Next j
   Next i
   Dim resultado As String
   For i = LBound(numeros) To UBound(numeros)
      If resultado <> "" Then resultado = resultado & "+"
      resultado = resultado & numeros(i)
   Next i
   OrdenarAscendente = resultado
End Function

'---------------------------------------------------------------------
' AYUDANTES GENÉRICOS
'---------------------------------------------------------------------

Private Function EnLista(ByVal valor As String, ParamArray listaValida() As Variant) As Boolean
    ' Un ParamArray no se puede pasar directamente a otra función;
    ' primero hay que copiarlo a un array normal.
    Dim copia() As Variant
    copia = listaValida
    EnLista = EnListaArray(valor, copia)
End Function

' Igual que EnLista pero recibe un array ya armado (para uso interno de EvalLista)
Private Function EnListaArray(ByVal valor As String, valores As Variant) As Boolean
    Dim i As Long
    For i = LBound(valores) To UBound(valores)
        If valor = CStr(valores(i)) Then
            EnListaArray = True
            Exit Function
        End If
    Next i
    EnListaArray = False
End Function

' ¿El valor (numérico, como texto) cae dentro de [minimo, maximo]?
Private Function EnRango(ByVal valorTexto As String, ByVal minimo As Double, ByVal maximo As Double) As Boolean
    If Not IsNumeric(valorTexto) Then
        EnRango = False
        Exit Function
    End If
    Dim valor As Double
    valor = CDbl(valorTexto)
    EnRango = (valor >= minimo And valor <= maximo)
End Function

' Texto legible con los valores válidos separados por " / "
Private Function UnirValores(valores As Variant) As String
   Dim i As Long
   Dim resultado As String
   For i = LBound(valores) To UBound(valores)
      If i > LBound(valores) Then resultado = resultado & " / "
      resultado = resultado & CStr(valores(i))
   Next i
   UnirValores = resultado
End Function
Private Sub EvalLista(ByVal ctrl As MSForms.Control, ByVal etiqueta As String, _
                       valoresValidos As Variant, ByRef contador As Long, ByRef mensaje As String)
    If Not EnListaArray(CStr(ctrl.Value), valoresValidos) Then
        contador = contador + 1
        mensaje = mensaje & "   - " & etiqueta & ": tienes '" & ctrl.Value & _
                  "', cámbialo por " & UnirValores(valoresValidos) & "." & vbNewLine
    End If
End Sub

Private Sub EvalRango(ByVal ctrl As MSForms.Control, ByVal etiqueta As String, _
                       ByVal minimo As Double, ByVal maximo As Double, ByRef contador As Long, ByRef mensaje As String)
    If Not EnRango(CStr(ctrl.Value), minimo, maximo) Then
        contador = contador + 1
        mensaje = mensaje & "   - " & etiqueta & ": tienes '" & ctrl.Value & _
                  "', debe estar entre " & minimo & " y " & maximo & "." & vbNewLine
    End If
End Sub

'---------------------------------------------------------------------
' CRITERIOS POR PRODUCTO (una función = una columna de la tabla)
'---------------------------------------------------------------------

Private Function CumpleGFlexBasic() As Boolean
    ' T22=Icc<10kA | T21=In<100A | T19=IP 31-66 | T28=Ingreso Inferior
    ' T12=Altura | T13=Ancho | T15=Profundidad | T112/T113=Base/Izaje No Aplica
    ' T114=Sujeción Pared/Poste
    CumpleGFlexBasic = _
        EnRango(Me.T22.Value, 0, 10) And _
        EnRango(Me.T21.Value, 0, 100) And _
        EnLista(Me.T19.Value, "31", "66") And _
        EnLista(Me.T28.Value, "Inferior") And _
        EnLista(Me.T12.Value, "600", "700", "800", "900") And _
        EnLista(Me.T13.Value, "400", "520", "620") And _
        EnLista(Me.T15.Value, "300") And _
        EnLista(Me.T112.Value, "No Aplica") And _
        EnLista(Me.T113.Value, "No Aplica") And _
        EnLista(Me.T114.Value, "Pared", "Poste")
End Function

Private Function CumpleGFlexNeo() As Boolean
    ' T22=Icc<25kA | T21=In<400A
    CumpleGFlexNeo = _
        EnRango(Me.T22.Value, 0, 25) And _
        EnRango(Me.T21.Value, 0, 400) And _
        EnLista(Me.T19.Value, "31", "42", "54", "55", "66","66") And _
        EnLista(Me.T28.Value, "Inferior") And _
        EnLista(Me.T12.Value, "600", "900", "1100") And _
        EnLista(Me.T13.Value, "400", "720", "1120") And _
        EnLista(Me.T15.Value, "400") And _
        EnLista(Me.T112.Value, "No Aplica") And _
        EnLista(Me.T113.Value, "No Aplica") And _
        EnLista(Me.T114.Value, "Pared", "Poste")
End Function

Private Function CumpleTPowerCompacto() As Boolean
    ' T22=Icc 25-40kA | T21=In 1250A
    CumpleTPowerCompacto = _
        EnRango(Me.T22.Value, 25, 40) And _
        EnRango(Me.T21.Value, 0, 1250) And _
        EnLista(Me.T19.Value, "30", "31", "42", "54","55") And _
        EnLista(Me.T28.Value, "Inferior", "Superior", "Superior / Inferior") And _
        EnLista(Me.T12.Value, "1600", "1800", "2000") And _
         EnRango(Me.T13.Value, 0, 4400) And _
        EnLista(Me.T15.Value, "400") And _
        EnLista(Me.T112.Value, "Zocalo") And _
        EnLista(Me.T113.Value, "Armella")
End Function

Private Function CumpleTPowerModular() As Boolean
    ' T22=Icc 40-65kA | T21=In 1600-4000A
    CumpleTPowerModular = _
        EnRango(Me.T22.Value, 40, 65) And _
        EnRango(Me.T21.Value, 1600, 4000) And _
        EnLista(Me.T19.Value, "31","42", "54","55") And _
        EnLista(Me.T28.Value, "Inferior", "Superior", "Superior / Inferior") And _
        EnLista(Me.T12.Value, "1800", "2000", "2200") And _
        EnRango(Me.T13.Value, 0, 4400)  And _
        EnLista(Me.T15.Value, "600", "800", "1000") And _
        EnLista(Me.T112.Value, "Zocalo", "Base estructural") And _
        EnLista(Me.T113.Value, "Armella", "Angulo de izaje")
End Function

Private Function CumpleTPowerHeavy() As Boolean
    ' T22=Icc>65kA | T21=In>4000A
    ' definido cómo se guarda ese dato en el formulario.
    CumpleTPowerHeavy = _
        EnRango(Me.T22.Value, 65, 999999) And _
        EnRango(Me.T21.Value, 4000, 9999999) And _
        EnLista(Me.T19.Value, "31","42", "54","55") And _
        EnLista(Me.T28.Value, "Inferior", "Superior", "Superior / Inferior") And _
        EnLista(Me.T12.Value, "2000", "2200") And _
         EnRango(Me.T13.Value, 0, 4400)  And _
        EnLista(Me.T15.Value, "1000") And _
        EnLista(Me.T112.Value, "Zocalo", "Base estructural") And _
        EnLista(Me.T113.Value, "Armella", "Angulo de izaje") 
End Function


'---------------------------------------------------------------------
' DIAGNÓSTICO POR PRODUCTO
' Cada función recorre los mismos criterios que su "Cumple..." gemela,
' pero en vez de detenerse en el primer False, cuenta cuántos criterios
' fallan y arma un mensaje con qué cambiar en cada uno.
' Devuelve: la cantidad de campos en conflicto (0 = sí cumple).
' Por referencia: el texto explicando cada conflicto.
'---------------------------------------------------------------------

Private Function DiferenciasGFlexBasic(ByRef mensaje As String) As Long
    Dim n As Long, msg As String
    EvalRango Me.T22, "Icc (T22)", 0, 10, n, msg
    EvalRango Me.T21, "In (T21)", 0, 100, n, msg
    EvalLista Me.T19, "IP (T19)", Array("31", "66"), n, msg
    EvalLista Me.T28, "Ingreso de acometidas (T28)", Array("Inferior"), n, msg
    EvalLista Me.T12, "Altura (T12)", Array("600", "700", "800", "900"), n, msg
    EvalLista Me.T13, "Ancho (T13)", Array("400", "520", "620"), n, msg
    EvalLista Me.T15, "Profundidad (T15)", Array("300"), n, msg
    EvalLista Me.T112, "Base (T112)", Array("No Aplica"), n, msg
    EvalLista Me.T113, "Izaje (T113)", Array("No Aplica"), n, msg
    EvalLista Me.T114, "Sujeción (T114)", Array("Pared", "Poste"), n, msg
    mensaje = msg
    DiferenciasGFlexBasic = n
End Function

Private Function DiferenciasGFlexNeo(ByRef mensaje As String) As Long
    Dim n As Long, msg As String
    EvalRango Me.T22, "Icc (T22)", 0, 25, n, msg
    EvalRango Me.T21, "In (T21)", 0, 400, n, msg
    EvalLista Me.T19, "IP (T19)", Array("31", "42", "54", "55", "66"), n, msg
    EvalLista Me.T28, "Ingreso de acometidas (T28)", Array("Inferior"), n, msg
    EvalLista Me.T12, "Altura (T12)", Array("600", "900", "1100"), n, msg
    EvalLista Me.T13, "Ancho (T13)", Array("400", "720", "1120"), n, msg
    EvalLista Me.T15, "Profundidad (T15)", Array("400"), n, msg
    EvalLista Me.T112, "Base (T112)", Array("No Aplica"), n, msg
    EvalLista Me.T113, "Izaje (T113)", Array("No Aplica"), n, msg
    EvalLista Me.T114, "Sujeción (T114)", Array("Pared", "Poste"), n, msg
    mensaje = msg
    DiferenciasGFlexNeo = n
End Function

Private Function DiferenciasTPowerCompacto(ByRef mensaje As String) As Long
    Dim n As Long, msg As String
    EvalRango Me.T22, "Icc (T22)", 25, 40, n, msg
    EvalRango Me.T21, "In (T21)", 0, 1250, n, msg
    EvalLista Me.T19, "IP (T19)", Array("30", "31", "42", "54", "55"), n, msg
    EvalLista Me.T28, "Ingreso de acometidas (T28)", Array("Inferior", "Superior", "Superior / Inferior"), n, msg
    EvalLista Me.T12, "Altura (T12)", Array("1600", "1800", "2000"), n, msg
    EvalRango Me.T13, "Ancho (T13)", 0, 4400, n, msg
    EvalLista Me.T15, "Profundidad (T15)", Array("400"), n, msg
    EvalLista Me.T112, "Base (T112)", Array("Zocalo"), n, msg
    EvalLista Me.T113, "Izaje (T113)", Array("Armella"), n, msg
    mensaje = msg
    DiferenciasTPowerCompacto = n
End Function

Private Function DiferenciasTPowerModular(ByRef mensaje As String) As Long
    Dim n As Long, msg As String
    EvalRango Me.T22, "Icc (T22)", 40, 65, n, msg
    EvalRango Me.T21, "In (T21)", 1600, 4000, n, msg
    EvalLista Me.T19, "IP (T19)", Array("31", "42", "54", "55"), n, msg
    EvalLista Me.T28, "Ingreso de acometidas (T28)", Array("Inferior", "Superior", "Superior / Inferior"), n, msg
    EvalLista Me.T12, "Altura (T12)", Array("1800", "2000", "2200"), n, msg
    EvalRango Me.T13, "Ancho (T13)", 0, 4400, n, msg
    EvalLista Me.T15, "Profundidad (T15)", Array("600", "800", "1000"), n, msg
    EvalLista Me.T112, "Base (T112)", Array("Zocalo", "Base estructural"), n, msg
    EvalLista Me.T113, "Izaje (T113)", Array("Armella", "Angulo de izaje"), n, msg
    mensaje = msg
    DiferenciasTPowerModular = n
End Function

Private Function DiferenciasTPowerHeavy(ByRef mensaje As String) As Long
    Dim n As Long, msg As String
    EvalRango Me.T22, "Icc (T22)", 65, 999999, n, msg
    EvalRango Me.T21, "In (T21)", 4000, 9999999, n, msg
    EvalLista Me.T19, "IP (T19)", Array("31", "42", "54", "55"), n, msg
    EvalLista Me.T28, "Ingreso de acometidas (T28)", Array("Inferior", "Superior", "Superior / Inferior"), n, msg
    EvalLista Me.T12, "Altura (T12)", Array("2000", "2200"), n, msg
    EvalRango Me.T13, "Ancho (T13)", 0, 4400, n, msg
    EvalLista Me.T15, "Profundidad (T15)", Array("1000"), n, msg
    EvalLista Me.T112, "Base (T112)", Array("Zocalo", "Base estructural"), n, msg
    EvalLista Me.T113, "Izaje (T113)", Array("Armella", "Angulo de izaje"), n, msg
    mensaje = msg
    DiferenciasTPowerHeavy = n
End Function


'---------------------------------------------------------------------
' ARMA EL MENSAJE DE SUGERENCIA CUANDO NINGÚN PRODUCTO CUMPLE DEL TODO
' Busca el/los producto(s) que necesitan MENOS cambios y los muestra.
'---------------------------------------------------------------------

Private Function ConstruirSugerencia() As String
    Const TOTAL_PRODUCTOS As Long = 5
    Dim nombres(1 To TOTAL_PRODUCTOS) As String
    Dim diffs(1 To TOTAL_PRODUCTOS) As Long
    Dim mensajes(1 To TOTAL_PRODUCTOS) As String

    nombres(1) = "G-Flex Basic"
    nombres(2) = "G-Flex NEO"
    nombres(3) = "T-Power Compacto"
    nombres(4) = "T-Power Modular"
    nombres(5) = "T-Power Heavy"

    diffs(1) = DiferenciasGFlexBasic(mensajes(1))
    diffs(2) = DiferenciasGFlexNeo(mensajes(2))
    diffs(3) = DiferenciasTPowerCompacto(mensajes(3))
    diffs(4) = DiferenciasTPowerModular(mensajes(4))
    diffs(5) = DiferenciasTPowerHeavy(mensajes(5))

    Dim minDiff As Long
    minDiff = diffs(1)
    Dim i As Long
    For i = 2 To TOTAL_PRODUCTOS
        If diffs(i) < minDiff Then minDiff = diffs(i)
    Next i

    Dim resultado As String
    resultado = "Con los datos actuales no se cumple ninguna configuración completa." & vbNewLine
    resultado = resultado & "La(s) opción(es) más cercana(s) necesita(n) " & minDiff & " cambio(s):" & vbNewLine & vbNewLine

    For i = 1 To TOTAL_PRODUCTOS
        If diffs(i) = minDiff Then
            resultado = resultado & "Para " & nombres(i) & ":" & vbNewLine & mensajes(i) & vbNewLine
        End If
    Next i

    ConstruirSugerencia = resultado
End Function


'---------------------------------------------------------------------
' MANEJO DE LOS OPTIONBUTTON (reset genérico + activar el par ganador)
'---------------------------------------------------------------------

Private Sub ResetOpciones()
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            ctrl.Value = False
            ctrl.Enabled = False
        End If
    Next ctrl
End Sub

Private Sub ActivarPar(ByVal opcion1 As MSForms.OptionButton, ByVal opcion2 As MSForms.OptionButton)
    opcion1.Enabled = True
    opcion1.Value = True
    opcion2.Enabled = True
    opcion2.Value = True
End Sub


'---------------------------------------------------------------------
' BOTÓN "VALIDAR"
'---------------------------------------------------------------------

Private Sub Validar_Click()
    ResetOpciones

    If CumpleGFlexBasic() Then
        ActivarPar Me.OptionGB, Me.OptionGB2
        Exit Sub

    ElseIf CumpleGFlexNeo() Then
        ActivarPar Me.OptionGN, Me.OptionGN2
        Exit Sub

    ElseIf CumpleTPowerCompacto() Then
        ActivarPar Me.OptionTC, Me.OptionTC2
        Exit Sub

    ElseIf CumpleTPowerModular() Then
        ActivarPar Me.OptionTM, Me.OptionTM2
        Exit Sub

    ElseIf CumpleTPowerHeavy() Then
        ActivarPar Me.OptionTH, Me.OptionTH2
        Exit Sub
    End If

    ' Ninguna configuración se cumple del todo: sugerir qué cambiar
    MsgBox ConstruirSugerencia(), vbInformation, "Ajustes sugeridos"
END Sub    
'Boton registar
Function ObtenerLineaSugerida() As String
    
   Dim lineaSugerida As String
   lineaSugerida = ""
   
   If OptionTH.Value = True Then
      lineaSugerida = "T-Power Heavy"
   ElseIf OptionTM.Value = True Then
      lineaSugerida = "T-Power Modular"
   ElseIf OptionTC.Value = True Then
      lineaSugerida = "T-Power Compacto"
   ElseIf OptionGN.Value = True Then
      lineaSugerida = "G-Flex NEO"
   ElseIf OptionGB.Value = True Then
      lineaSugerida = "G-Flex Basic"
   ElseIf OptionBST.Value = True Then
      lineaSugerida = "BlokSeT"
   End If
   
   ObtenerLineaSugerida = lineaSugerida

End Function
Private Sub Registrar_Click()
   with thisWorkbook.sheets("Portada_Mestra")
      ' c3 a c68
      '==== 0. ALCANCE PROYECTO ====
         .Range("C3").Value = Me.T00.Value
         .range("C4").Value = Me.T01.Value
         .Range("C5").Value = Me.T02.Value
         .Range("C6").Value = Me.T03.Value
         .Range("C7").Value = Me.T04.Value
         .Range("C8").Value = Me.T05.Value
         .Range("C9").Value = Me.T06.Value
         .Range("C10").Value = Me.T07.Value
         .Range("C11").Value = Me.T08.Value
      ' ==== 1. CARACTERISTICAS METALMECANICAS ====
         .Range("C13").Value = Me.T11.Value
         .Range("C14").Value = Me.T12.Value & "x" & Me.T13.Value & "x" & Me.T15.Value
         .Range("C15").Value = Me.T16.Value
         .Range("C16").Value = Me.T17.Value
         .Range("C17").Value = Me.T18.Value
         .Range("C18").Value = Me.T19.Value
         .Range("C19").Value = Me.T110.Value
         .Range("C20").Value = Me.T111.Value
         .Range("C21").Value = Me.T112.Value
         .Range("C22").Value = Me.T113.Value
         .Range("C23").Value = Me.T114.Value
      '==== 2. CARACTERISTICAS DE ELECTRICIDAD ====
         .Range("C25").Value = Me.T21.Value
         .Range("C26").Value = Me.T22.Value
         .Range("C27").Value = Me.T23.Value
         .Range("C28").Value = Me.T24.Value
         .Range("C29").Value = Me.T25.Value
         .Range("C30").Value = Me.T26.Value
         .Range("C31").Value = Me.T27.Value
         .Range("C32").Value = Me.T28.Value
         .Range("C33").Value = Me.T29.Value
         .Range("C34").Value = Me.T210.Value
      '==== 3. CARACTERISTICAS DE POTENCIA ====   
         .Range("C36").Value = Me.T31.Value
         .Range("C37").Value = Me.T32.Value
         .Range("C38").Value = Me.T33.Value
         .Range("C39").Value = Me.T34.Value
         .Range("C40").Value = Me.T35.Value
         .Range("C41").Value = Me.T36.Value
         .Range("C42").Value = Me.T37.Value
         .Range("C44").Value = Me.L1.Caption
         .Range("C45").Value = Me.L2.Caption
         .Range("C46").Value = Me.L3.Caption
         .Range("C47").Value = Me.L4.Caption
         .Range("C48").Value = Me.L5.Caption
      '==== 4. CARACTERISTICAS DE SCONTROL ====
         .Range("C50").Value = Me.T41.Value
         .Range("C51").Value = Me.T42.Value
         .Range("C52").Value = Me.T43.Value
         .Range("C53").Value = Me.T44.Value
         .Range("C54").Value = Me.T45.Value
         .Range("C55").Value = Me.T46.Value
         .Range("C56").Value = Me.T47.Value
         .Range("C57").Value = Me.T48.Value
      '==== 5. CARACTERISTICAS DE COMUNICACION ====   
         .Range("C59").Value = Me.T51.Value
      '==== 6. CARACTERISTICAS DE EMBALAJE ====   
         .Range("C61").Value = Me.T61.Value
      '==== 7. PAQUETE DE PRUEBAS ====   
         .Range("C63").Value = Me.T71.Value
      '==== 8. IDENTIFICACION Y ETIQUETADO ==== 
      .Range("C65").Value = Me.T81.Value
      .Range("C66").Value = Me.T82.Value
      .Range("C67").Value = Me.T83.Value
      .Range("C68").Value = Me.T84.Value
      .Range("C69").Value = Me.T85.Value
      .Range("C70").Value = Me.T86.Value
      .Range("C71").Value = Me.T87.Value
      .Range("C72").Value = Me.T88.Value
      .Range("C73").Value = Me.T89.Value
      .Range("C74").Value = Me.T810.Value
   END WITH
   MSGBOX "El formulario se ha guardado correctamente.", vbInformation   


   ' PREGUNTA DE PARA QUIÉN VA DIRIGIDO EL PAQUETE
   Dim respuesta As String
   respuesta = InputBox("¿Para quién va dirigido el paquete?" & vbNewLine & _
                        "1 = Pruebas FAT" & vbNewLine & _
                        "2 = APE" & vbNewLine & _
                        "3 = Cliente" & vbNewLine & _
                        "4 = Todas", "Información")

   Select Case respuesta
      Case "1"
         GenerarPaquete 1, Me.T11.Value
      Case "2"
         GenerarPaquete 2, Me.T11.Value
      Case "3"
         GenerarPaquete 3, Me.T11.Value
      Case "4"
         GenerarPaquete 4, Me.T11.Value
      Case Else
         MsgBox "Opción no válida. Ingresa un número del 1 al 4.", vbExclamation, "Información"
   End Select
   '________________________________________________________________________
   Dim lineaSugerida As String
    lineaSugerida = ObtenerLineaSugerida()
    
    If lineaSugerida = "" Then
        MsgBox "No se ha seleccionado ninguna línea de producto.", vbExclamation
        Exit Sub
    End If

   ' Preguntar si desea ir a diseño
   If MsgBox("¿Desea ir a la pantalla de diseño?", vbQuestion + vbYesNo, "Información") = vbYes Then
      Call AbrirLineaYCopiarPortada(lineaSugerida)
      Unload Me
   Else
      Unload Me
   End If

End Sub