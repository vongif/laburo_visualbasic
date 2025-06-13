VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZREC 
   Caption         =   "ZREC"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12585
   OleObjectBlob   =   "anulacion_masiva.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numeroPedidoActual As Long
Dim posicionActual As Long
Dim datosTemp As Collection
Dim marcasPedidoActual As String ' NUEVO
Dim desactivandoCodigo As Boolean




Private Sub UserForm_Initialize()
    If numeroPedidoActual = 0 Then numeroPedidoActual = 1
    posicionActual = 10
    Set datosTemp = New Collection

    With lstResumen
    .ColumnCount = 7
    .ColumnWidths = "75 pt;80 pt;100 pt;70 pt;55 pt;60 pt;100 pt"
    End With

    txtFecha.SetFocus
End Sub




Private Sub txtCodigo_Change()

    If desactivado Then Exit Sub

    Dim codigoParcial As String
    Dim hojaCodigos As Worksheet
    Dim celda As Range
    Dim Descripcion As String
    Dim sabor As String
    Dim rangoBusqueda As Range
    Dim encontrado As Boolean

    Set hojaCodigos = ThisWorkbook.Sheets("codigos")
    codigoParcial = Trim(Me.txtCodigo.Text)
    
    If codigoParcial = "" Then
        desactivado = True
        Me.txtDescripcion.Value = ""
        desactivado = False
        Exit Sub
    End If

    Set rangoBusqueda = hojaCodigos.Range("A2:A3402")
    encontrado = False

    For Each celda In rangoBusqueda
        If LCase(Left(celda.Value, Len(codigoParcial))) = LCase(codigoParcial) Then
            Descripcion = hojaCodigos.Cells(celda.Row, 2).Value
            sabor = hojaCodigos.Cells(celda.Row, 3).Value
            desactivado = True
            Me.txtDescripcion.Value = Descripcion & " - " & sabor
            desactivado = False
            encontrado = True
            Exit For
        End If
    Next celda

    If Not encontrado Then
        desactivado = True
        Me.txtDescripcion.Value = ""
        desactivado = False
    End If

End Sub



Private Sub txtCodigo_AfterUpdate()
    If desactivandoCodigo Then Exit Sub

    Dim base As String
    base = Trim(txtCodigo.Text)

    Dim largo As Integer
    largo = Len(base)

    If largo >= 6 Then Exit Sub

    Dim totalFinal As Integer
    Dim cerosFaltantes As Integer

    If largo < 4 Then
        totalFinal = 5
    Else
        totalFinal = 6
    End If

    cerosFaltantes = totalFinal - largo

    If cerosFaltantes > 0 Then
        txtCodigo.Text = base & String(cerosFaltantes, "0")
        txtCodigo.SelStart = Len(base)
        txtCodigo.SelLength = cerosFaltantes
    End If
End Sub



Private Sub txtDescripcion_Change()
    Dim hojaCodigos As Worksheet
    Dim celda As Range
    Dim texto As String
    Dim coincidencias As Collection
    Dim Descripcion As String, sabor As String

    Set hojaCodigos = ThisWorkbook.Sheets("codigos")
    texto = LCase(Trim(Me.txtDescripcion.Text))

    lstSugerencias.Clear
    lstSugerencias.Visible = False

    If texto = "" Then Exit Sub

    For Each celda In hojaCodigos.Range("B2:B3402")
        Descripcion = LCase(celda.Value)
        sabor = hojaCodigos.Cells(celda.Row, 3).Value

        If InStr(Descripcion, texto) > 0 Then
            On Error Resume Next
            lstSugerencias.AddItem celda.Value & " - " & sabor
            lstSugerencias.List(lstSugerencias.ListCount - 1, 1) = celda.Row
            On Error GoTo 0
        End If
    Next celda

    If lstSugerencias.ListCount > 0 Then
        With lstSugerencias
            .Top = txtDescripcion.Top + txtDescripcion.Height
            .Left = txtDescripcion.Left
            .Width = txtDescripcion.Width
            .Height = 60
            .Visible = True
        End With
    End If
End Sub


Private Sub lstSugerencias_Click()
    Dim fila As Long
    Dim hojaCodigos As Worksheet
    Set hojaCodigos = ThisWorkbook.Sheets("codigos")

    ' Obtener la fila guardada en la segunda columna oculta de la lista
    fila = CLng(lstSugerencias.List(lstSugerencias.ListIndex, 1))

    ' Completar los campos
    txtDescripcion.Text = hojaCodigos.Cells(fila, 2).Value & " - " & hojaCodigos.Cells(fila, 3).Value
    txtCodigo.Text = hojaCodigos.Cells(fila, 1).Value

    ' Ocultar la lista
    lstSugerencias.Visible = False
End Sub



Private Sub txtCliente_Change()
    Dim hojaClientes As Worksheet
    Dim celda As Range
    Dim numeroCliente As String

    numeroCliente = Trim(Me.txtCliente.Text)
    Set hojaClientes = ThisWorkbook.Sheets("clientes")

    ' Limpiar el txtRepartos por defecto
    txtRepartos.Value = ""

    If Len(numeroCliente) <> 8 Or Not IsNumeric(numeroCliente) Then Exit Sub

    For Each celda In hojaClientes.Range("A2:A" & hojaClientes.Cells(hojaClientes.Rows.Count, "A").End(xlUp).Row)
        If Trim(celda.Value) = numeroCliente Then
            txtRepartos.Value = hojaClientes.Cells(celda.Row, 2).Value
            Exit For
        End If
    Next celda
End Sub




Private Sub txtDescripcion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lstSugerencias.Visible And lstSugerencias.ListCount > 0 Then
        Select Case KeyCode
            Case vbKeyDown
                If lstSugerencias.ListIndex < lstSugerencias.ListCount - 1 Then
                    lstSugerencias.ListIndex = lstSugerencias.ListIndex + 1
                ElseIf lstSugerencias.ListIndex = -1 Then
                    lstSugerencias.ListIndex = 0
                End If
                KeyCode = 0

            Case vbKeyUp
                If lstSugerencias.ListIndex > 0 Then
                    lstSugerencias.ListIndex = lstSugerencias.ListIndex - 1
                End If
                KeyCode = 0

            Case vbKeyReturn ' ENTER
                If lstSugerencias.ListIndex >= 0 Then
                    Call lstSugerencias_Click
                    KeyCode = 0
                End If

            Case vbKeyEscape
                lstSugerencias.Visible = False
                KeyCode = 0
        End Select
    End If
End Sub






Private Sub btnAgregarOtro_Click()
    If Not ValidarCampos() Then Exit Sub
       
      
    Dim tipoRecibo As String
    If btnR01.Value Then
    tipoRecibo = "R01"
    ElseIf btnR02.Value Then
    tipoRecibo = "R02"
    End If
    
    
    ' Verificar código duplicado
    Dim itemExistente As Collection
    Dim codigoIngresado As String
    codigoIngresado = txtCodigo.Value

    Dim i As Long
    For i = 1 To datosTemp.Count
        Set itemExistente = datosTemp(i)
        If itemExistente(6) = codigoIngresado Then
            MsgBox "Ya ingresaste este código en el pedido actual.", vbExclamation
            txtCodigo.SetFocus
            Exit Sub
        End If
    Next i
    
    
    
     ' Guardar los datos
    Dim item As Collection
    Set item = New Collection
    
    item.Add numeroPedidoActual                            ' 1 - N° Pedido (ID)
    item.Add tipoRecibo                                    ' 2 - Tipo de recibo
    item.Add Format(CDate(txtFecha.Value), "dd/mm/yyyy")   ' 3 - Fecha
    item.Add txtCliente.Value                              ' 4 - Cliente
    item.Add txtReferencia.Value                           ' 5 - Referencia
    item.Add txtCodigo.Value                               ' 6 - Código
    item.Add txtCantidad.Value                             ' 7 - Cantidad
    item.Add ObtenerMarcasSeleccionadas()
    item.Add chkGuiaAparte.Value                           ' 8 - Guía
    item.Add posicionActual                                ' 9 - Posición
    
    datosTemp.Add item
    MostrarItemEnLista item

    ' Preparar siguiente entrada
    posicionActual = posicionActual + 10
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    chkGuiaAparte.Value = False
    txtCodigo.SetFocus
End Sub

Private Function ValidarCampos() As Boolean
    ValidarCampos = False

    If Len(txtFecha.Value) <> 10 Or Not IsDate(txtFecha.Value) Then
    MsgBox "Ingresá una fecha válida en formato DD/MM/AAAA.", vbExclamation
    txtFecha.SetFocus: Exit Function
    End If

    If txtCliente.Value = "" Or txtCodigo.Value = "" Or txtCantidad.Value = "" Then
        MsgBox "Completá todos los campos antes de agregar.", vbExclamation
        Exit Function
    End If
    
    If txtReferencia.Value = "" Then
       MsgBox "Ingresá una referencia.", vbExclamation
       txtReferencia.SetFocus: Exit Function
    End If

    If Not IsNumeric(txtCliente.Value) Or Len(txtCliente.Value) <> 8 Then
        MsgBox "El cliente debe contener exactamente 8 números.", vbExclamation
        txtCliente.SetFocus: Exit Function
    End If

    'If Not IsNumeric(txtCodigo.Value) Or Len(txtCodigo.Value) < 5 Or Len(txtCodigo.Value) > 6 Then
    'MsgBox "El código debe contener 5 o 6 números.", vbExclamation
    'txtCodigo.SetFocus
    'Exit Function
    'End If

    If Not IsNumeric(txtCantidad.Value) Then
        MsgBox "La cantidad debe ser un número.", vbExclamation
        txtCantidad.SetFocus: Exit Function
    End If
    
    If Not btnR01.Value And Not btnR02.Value Then
    MsgBox "Seleccioná el tipo de recibo: R01 o R02.", vbExclamation
    btnR01.SetFocus
    Exit Function
    End If
    
    If Not (chkMastellone.Value Or chkDanone.Value Or chkNutricia.Value Or _
            chkCalsa.Value Or chkLario.Value Or chkLogistica.Value) Then
        MsgBox "Debés seleccionar al menos una empresa (marca).", vbExclamation
        Exit Function
    End If
    

    ValidarCampos = True
End Function


Private Function ObtenerMarcasSeleccionadas() As String
    Dim marcas As String
    If chkMastellone.Value Then marcas = marcas & "7199, "
    If chkDanone.Value Then marcas = marcas & "7100, "
    If chkNutricia.Value Then marcas = marcas & "5770, "
    If chkCalsa.Value Then marcas = marcas & "9001, "
    If chkLario.Value Then marcas = marcas & "9002, "
    If chkLogistica.Value Then marcas = marcas & "7140, "

    If Right(marcas, 2) = ", " Then marcas = Left(marcas, Len(marcas) - 2)
    ObtenerMarcasSeleccionadas = marcas
End Function



Private Sub MostrarItemEnLista(item As Collection)
    lstResumen.AddItem
    With lstResumen
    .List(.ListCount - 1, 0) = item(3) ' Fecha
    .List(.ListCount - 1, 1) = item(4) ' Cliente
    .List(.ListCount - 1, 2) = item(5) ' Referencia
    .List(.ListCount - 1, 3) = item(6) ' Código (Material)
    .List(.ListCount - 1, 4) = item(7) ' Cantidad
    .List(.ListCount - 1, 5) = item(2) ' Tipo de recibo
    .List(.ListCount - 1, 6) = item(8) ' Marcas / Org
    
    End With
End Sub


Private Sub btnConfirmar_Click()
    If datosTemp.Count = 0 Then
        ' Validación alternativa cuando no hay ítems
        If Len(txtFecha.Value) <> 10 Or Not IsDate(txtFecha.Value) Then
            MsgBox "Ingresá una fecha válida en formato DD/MM/AAAA.", vbExclamation
            txtFecha.SetFocus: Exit Sub
        End If

        If txtCliente.Value = "" Or Not IsNumeric(txtCliente.Value) Or Len(txtCliente.Value) <> 8 Then
            MsgBox "El cliente debe contener exactamente 8 números.", vbExclamation
            txtCliente.SetFocus: Exit Sub
        End If

        If txtReferencia.Value = "" Then
            MsgBox "Ingresá una referencia.", vbExclamation
            txtReferencia.SetFocus: Exit Sub
        End If

        If Not btnR02.Value Then
            MsgBox "Solo podés confirmar sin ítems si el recibo es R02.", vbExclamation
            btnR02.SetFocus: Exit Sub
        End If

        If Not chkLogistica.Value Then
            MsgBox "Solo podés confirmar sin ítems si la organización es Logística.", vbExclamation
            chkLogistica.SetFocus: Exit Sub
        End If

        If MsgBox("Estás por confirmar un pedido sin ítems. ¿Querés continuar?", vbYesNo + vbQuestion, "Confirmar sin ítems") = vbNo Then
            Exit Sub
        End If
    End If

    Call GuardarYLimpiar
    MsgBox "Pedido confirmado y cargado en la planilla.", vbInformation
End Sub




Private Sub btnEliminar_Click()
    Dim idx As Long: idx = lstResumen.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná un pedido para eliminar.", vbExclamation
        Exit Sub
    End If

    datosTemp.Remove idx + 1
    Call RecalcularPosiciones

    MsgBox "Ítem eliminado y posiciones actualizadas.", vbInformation
End Sub

Private Sub btnModificar_Click()
    Dim idx As Long: idx = lstResumen.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná un pedido para modificar.", vbExclamation
        Exit Sub
    End If

    Dim item As Collection: Set item = datosTemp(idx + 1)

    txtFecha.Value = item(3)
    txtCliente.Value = item(4)
    txtReferencia.Value = item(5)
    txtCodigo.Value = item(6)
    txtCantidad.Value = item(7)
    chkGuiaAparte.Value = item(9)
    Call RestaurarMarcas(item(8))
    datosTemp.Remove idx + 1
    lstResumen.RemoveItem idx

    Call RecalcularPosiciones
    txtCodigo.SetFocus
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub LimpiarTodo()
    Set datosTemp = New Collection
    posicionActual = 10
    txtFecha.Value = ""
    txtCliente.Value = ""
    txtReferencia.Value = ""
    desactivandoCodigo = True
    txtCodigo.Value = ""
    desactivandoCodigo = False
    txtCantidad.Value = ""
    chkGuiaAparte.Value = False
    btnR01.Value = False
    btnR02.Value = False
    lstResumen.Clear
    marcasPedidoActual = ""
    txtRepartos.Value = ""
    Call LimpiarChecks
    txtFecha.SetFocus
End Sub

Private Sub LimpiarChecks()
    chkMastellone.Value = False
    chkDanone.Value = False
    chkNutricia.Value = False
    chkCalsa.Value = False
    chkLario.Value = False
    chkLogistica.Value = False
End Sub

Private Sub RestaurarMarcas(marcas As String)
    LimpiarChecks
    If InStr(marcas, "7199") Then chkMastellone.Value = True
    If InStr(marcas, "7100") Then chkDanone.Value = True
    If InStr(marcas, "5770") Then chkNutricia.Value = True
    If InStr(marcas, "9001") Then chkCalsa.Value = True
    If InStr(marcas, "9002") Then chkLario.Value = True
    If InStr(marcas, "7140") Then chkLogistica.Value = True
End Sub



'==================
'CHECK COLOR & UN SOLO CHECK
'==================

Private Sub CambiarColorOptionSeleccionado()
    Dim ctrl As Control

    For Each ctrl In Me.frmMarcas.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Value = True Then
                ctrl.BackColor = RGB(173, 216, 230) ' Azul claro
            Else
                ctrl.BackColor = &H8000000A  ' Color por defecto
            End If
        End If
    Next ctrl
End Sub

Private Sub chkMastellone_Change()
    CambiarColorOptionSeleccionado
End Sub

Private Sub chkDanone_Change()
    CambiarColorOptionSeleccionado
End Sub

Private Sub chkNutricia_Change()
    CambiarColorOptionSeleccionado
End Sub

Private Sub chkCalsa_Change()
    CambiarColorOptionSeleccionado
End Sub

Private Sub chkLario_Change()
    CambiarColorOptionSeleccionado
End Sub

Private Sub chkLogistica_Change()
    CambiarColorOptionSeleccionado
End Sub





Private Sub RecalcularPosiciones()
    Dim i As Long, item As Collection
    Dim nuevaPosicion As Long: nuevaPosicion = 10

    lstResumen.Clear

    For i = 1 To datosTemp.Count
        Set item = datosTemp(i)
        If item.Count = 6 Then item.Remove 6
        item.Add nuevaPosicion
        nuevaPosicion = nuevaPosicion + 10
        MostrarItemEnLista item
    Next i

    posicionActual = nuevaPosicion
End Sub

Private Sub btnNuevoCliente_Click()
    
    If datosTemp.Count = 0 Then
        MsgBox "No hay ítems para guardar.", vbExclamation
        Exit Sub
    End If

    Call GuardarYLimpiar
    
    
End Sub




' ===========================
' NAVEGACIÓN CON ENTER
' ===========================
Private Sub txtFecha_Change()
    Dim txt As String
    Dim partes() As String
    Static editando As Boolean
    If editando Then Exit Sub
    
    ' Si se usan barras manualmente…
    If InStr(txtFecha.Text, "/") > 0 Then
        ' Dividimos la cadena en partes
        partes = Split(txtFecha.Text, "/")
        ' Si hay dos partes y el segundo contiene más de dos dígitos,
        ' significa que se está ingresando el año.
        If UBound(partes) = 1 And Len(partes(1)) > 2 Then
            editando = True
            ' Formateamos: Primer parte (día), segundo parte (mes) y lo que resta como año.
            txtFecha.Text = partes(0) & "/" & Left(partes(1), 2) & "/" & Mid(partes(1), 3)
            txtFecha.SelStart = Len(txtFecha.Text)
            editando = False
        End If
        Exit Sub
    End If

    ' Si no se usaron barras, procesamos el ingreso como cadena de solo números.
    txt = Replace(txtFecha.Text, "/", "")
    If Not IsNumeric(txt) Then Exit Sub
    If Len(txt) > 8 Then Exit Sub  ' Máximo 8 dígitos (DDMMAAAA)
    
    editando = True
    Select Case Len(txt)
        Case 1 To 2
            txtFecha.Text = txt
        Case 3 To 4
            ' Se asume que son día y mes sin separador: "DDMM" ? "DD/MM"
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3)
        Case 5 To 6
            ' "DDMMAA" ? "DD/MM/AA" (aunque falte año completo)
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2)
        Case 7 To 8
            ' Formato completo: "DDMMAAAA" ? "DD/MM/AAAA"
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2) & "/" & Mid(txt, 5)
    End Select
    txtFecha.SelStart = Len(txtFecha.Text)
    editando = False
End Sub

Private Sub txtFecha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim partes() As String
        partes = Split(txtFecha.Text, "/")
        
        ' Si sólo se tiene día y mes, se completa con el año actual.
        If UBound(partes) = 1 Then
            txtFecha.Text = txtFecha.Text & "/" & Year(Date)
        End If
        
        ' Validar que la fecha sea real.
        If Not IsDate(txtFecha.Text) Then
            MsgBox "La fecha ingresada no es válida.", vbExclamation, "Error de fecha"
            txtFecha.Text = ""
            Exit Sub
        End If
        
        ' Enfocar el siguiente control.
        txtCliente.SetFocus
    End If
End Sub




Private Sub txtCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtReferencia.SetFocus
End Sub

Private Sub txtReferencia_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCodigo.SetFocus
End Sub


Private Sub txtCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCantidad.SetFocus
End Sub

Private Sub txtCantidad_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnAgregarOtro.SetFocus
End Sub


Private Sub GuardarYLimpiar()
    Dim hoja As Worksheet: Set hoja = ThisWorkbook.Sheets("ZREC")
    Dim fila As Long
    fila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row
    If fila < 2 Then fila = 2 Else fila = fila + 1

    Const colPedido As String = "A"
    Const colCliente As String = "B"
    Const colGuia As String = "D"
    Const colFecha As String = "J"
    Const colCodigo As String = "L"
    Const colCantidad As String = "O"
    Const colMarcas As String = "E"
    Const colPosicion As String = "K"
    Const colReferencia As String = "T"
    Const colTipoRecibo As String = "I"

    Dim hayGuiaAparte As Boolean
    Dim item As Collection

    ' Verificar si hay alguna guía aparte
    For Each item In datosTemp
        If item(5) = True Then
            hayGuiaAparte = True
            Exit For
        End If
    Next item

    ' Escribir los datos
    If datosTemp.Count > 0 Then
        For Each item In datosTemp
            hoja.Cells(fila, colPedido).Value = item(1)
            hoja.Cells(fila, colTipoRecibo).Value = item(2)
            hoja.Cells(fila, colFecha).Value = Format(CDate(item(3)), "yyyymmdd")
            hoja.Cells(fila, colCliente).Value = item(4)
            hoja.Cells(fila, colReferencia).Value = item(5)
            hoja.Cells(fila, colCodigo).Value = item(6)
            hoja.Cells(fila, colCantidad).Value = item(7)
            hoja.Cells(fila, colMarcas).Value = item(8) ' línea que agrega las marcas
            hoja.Cells(fila, colPosicion).Value = item(10)
            fila = fila + 1
        Next item
    Else
        ' Si no hay ítems, guardar solo los campos básicos
        hoja.Cells(fila, colPedido).Value = numeroPedidoActual
        hoja.Cells(fila, colTipoRecibo).Value = "R02"
        hoja.Cells(fila, colFecha).Value = Format(CDate(txtFecha.Value), "yyyymmdd")
        hoja.Cells(fila, colCliente).Value = txtCliente.Value
        hoja.Cells(fila, colReferencia).Value = txtReferencia.Value
        hoja.Cells(fila, colMarcas).Value = "7140"
        ' Los demás campos se dejan vacíos
    End If

    numeroPedidoActual = numeroPedidoActual + 1
    Call LimpiarTodo
End Sub


