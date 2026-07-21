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
      with Me.T07
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
private sub T11_Change()
   dim configuracion as string
   configuracion = Me.T11.Value
   with Me.T12
      .Clear
      select case configuracion
         case "T-Power"
            .AddItem "1600"
            .AddItem "1800"
            .AddItem "2000"
            .AddItem "2200"
         case "G-Flex"
            .AddItem "600"
            .AddItem "700"
            .AddItem "800"
            .AddItem "900"
            .AddItem "1100"
         case "BlokSeT"
            .AddItem "por definir"
         case "D-Critical"
            .AddItem "por definir"
         case "S-Energy"
            .AddItem "por definir"
         case "M-Board"
            .AddItem "por definir"
         case "G-Sync"
            .AddItem "por definir"
         case "A-Tranfer"
            .AddItem "por definir"
      end select
   end with
   with Me.T15
      .Clear
      select case configuracion
         case "T-Power"
            .AddItem "400"
            .AddItem "600"
            .AddItem "800"
            .AddItem "1000"
         case "G-Flex"
            .AddItem "300"
            .AddItem "400"
         case "BlokSeT"
            .AddItem "por definir"
         case "D-Critical"
            .AddItem "por definir"
         case "S-Energy"
            .AddItem "por definir"
         case "M-Board"
            .AddItem "por definir"
         case "G-Sync"
            .AddItem "por definir"
         case "A-Tranfer"
            .AddItem "por definir"
      end select
   end with

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
End sub

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
        Case "T-Power"
            medidas = Array(1120, 720, 400, 200) ' orden descendente
        Case "G-Flex"
            medidas = Array(1120, 720, 620, 520, 400) ' orden descendente
        Case Else
            T14.Value = "Seleccione un tipo válido (T-Power o G-Flex)"
            Exit Sub
    End Select

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

    Dim i As Long
    resultado = ""
    For i = 1 To combinaciones.Count
        If resultado <> "" Then resultado = resultado & vbNewLine
        resultado = resultado & combinaciones(i)
    Next i

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