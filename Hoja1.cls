'Hola panel
'======== DOCUMENTACION =========
Public Sub Worksheet_Activate()
    With Me.OLEObjects("portada").Object
        If .ListCount = 0 Then
        .AddItem "Añadir" 'H_Portada
        .AddItem "No anadir"
        End If
    End With
    With Me.OLEObjects("D_generales").Object
        .Clear
        .AddItem ""
        .AddItem "Hoja caracteristicas electricas APE" 'HC_APE
        .AddItem "Hoja caracteristicas electricas Cliente" 'HC_CLIENTE
        .AddItem "Hoja caracteristicas electricas FAT" 'HC_FAT
        .AddItem "Añadir todas las hojas"
    End With
    With Me.OLEObjects("D_generales2").Object
        .Clear
        .AddItem ""
        .AddItem "Hoja caracteristicas electricas APE" 'HC_APE
        .AddItem "Hoja caracteristicas electricas Cliente" 'HC_CLIENTE
        .AddItem "Hoja caracteristicas electricas FAT" 'HC_FAT
    End With
    With Me.OLEObjects("Simbologia").Object
        .Clear
        .AddItem "Simbologia electrica 1 y 2" 'SIM_2 'SIM_3
        .AddItem "Simbologia electrica 2 y 3" 'SIM_3 'SIM_4
        .AddItem "Simbologia electrica  y control 3 y 4" 'SIM_4 'SIM_5
        .AddItem "añadir todas las simbologias"
        .AddItem "No anadir"'
    End With
    With Me.OLEObjects("Placa").Object
        .Clear
        .AddItem "Placa de datos CTS" 'PLG_DATOS
        .AddItem "No anadir" 
    End With
    with Me.OLEObjects("Cajetines").Object
        .Clear
        .AddItem "Cajetin Mecanico, Unifilar, control" 'CajG_mecanico , CajG_unifilar, CajG_Control
        .AddItem "No anadir"
    End With

    CargarCombosLinea
    'ADICIONALES 
    With Me.OLEObjects("Simbologia2").Object
        .Clear
        .AddItem "Simbologia electrica 1" 'SIM_2
        .AddItem "Simbologia electrica 2" 'SIM_3
        .AddItem "Simbologia electrica 3" 'SIM_4
        .AddItem "Simbologia electrica 4" 'SIM_5
        .AddItem "No anadir"
    End With
     with Me.OLEObjects("Cajetines2").Object
        .Clear
        .AddItem "Cajetin Mecanico" 'CajG_mecanico
        .AddItem "Cajetin Unifilar" 'CajG_unifilar
        .AddItem "Cajetin Control" 'CajG_Control
        .AddItem "No anadir"
    End With
    With Me.OLEObjects("Indice").Object
        .Clear
        .AddItem "Añadir"
        .AddItem "No anadir"
    End With
End Sub
Public Sub CargarCombosLinea()
    Dim p As String, e As String, a As String

    Select Case LineaActual()
        Case "L_1": p = "Puertas TPOWER":   e = "Etiquetas TPOWER":   a = "Detalle de anclaje Tablero T-Power NEO"
        Case "L_2": p = "Puertas GFLEX_N":  e = "Etiquetas GFLEX_N":  a = "Detalle de anclaje Tablero GFLEX"
        Case "L_3": p = "Puertas GFLEX_B":  e = "Etiquetas GFLEX_B":  a = "Detalle de anclaje Tablero GFLEX"
        Case "L_4": p = "Puertas Blokset":  e = "Etiquetas Blokset":  a = "Detalle de anclaje Tablero Blokset"
    End Select

    LlenarCombo "puertas", p
    LlenarCombo "etiquetas", e
    LlenarCombo "anclaje", a
End Sub

Private Sub LlenarCombo(ByVal nombre As String, ByVal opcion As String)
    With Me.OLEObjects(nombre).Object
        .Clear
        If opcion <> "" Then
            .AddItem opcion
            .AddItem "No anadir"
        End If
    End With
End Sub

