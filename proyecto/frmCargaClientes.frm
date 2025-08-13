VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCargaClientes 
   Caption         =   "ZPVA"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9570.001
   OleObjectBlob   =   "frmCargaClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCargaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numeroPedidoActual As Long
Dim posicionActual As Long
Dim datosTemp As Collection

Private Sub UserForm_Initialize()
    numeroPedidoActual = ObtenerUltimoPedido() + 1
    posicionActual = 10
    Set datosTemp = New Collection

    With lstResumen
        .ColumnCount = 5
        .ColumnWidths = "80 pt;100 pt;80 pt;60 pt;60 pt"
    End With

    btnModificar.Enabled = False
    btnEliminar.Enabled = False

    txtFecha.SetFocus
End Sub

Private Function ObtenerUltimoPedido() As Long
    Dim hoja As Worksheet
    Set hoja = ThisWorkbook.Sheets("ZPVA")
    
    Dim ultimaFila As Long
    ultimaFila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 2 Then
        ObtenerUltimoPedido = 0
    Else
        Do While ultimaFila >= 2
            If IsNumeric(hoja.Cells(ultimaFila, "A").Value) Then
                ObtenerUltimoPedido = hoja.Cells(ultimaFila, "A").Value
                Exit Function
            End If
            ultimaFila = ultimaFila - 1
        Loop
        ObtenerUltimoPedido = 0
    End If
End Function

Private Sub btnAgregarOtro_Click()
    If Not ValidarCampos() Then Exit Sub

    If Left(txtCantidad.Text, 1) = "0" And Len(txtCantidad.Text) > 1 Then
        MsgBox "La cantidad no puede comenzar con cero.", vbExclamation
        txtCantidad.SetFocus
        Exit Sub
    End If

    Dim item As Collection
    Set item = New Collection

    item.Add Format(CDate(txtFecha.Value), "dd/mm/yyyy")
    item.Add txtCliente.Value
    item.Add txtCodigo.Value
    item.Add txtCantidad.Value
    item.Add chkGuiaAparte.Value
    item.Add posicionActual

    datosTemp.Add item
    MostrarItemEnLista item

    posicionActual = posicionActual + 10
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    txtCodigo.SetFocus

    ActualizarEstadoBotones
End Sub

Private Sub chkGuiaAparte_Click()
    If chkGuiaAparte.Value = True Then
        chkGuiaAparte.BackColor = RGB(0, 250, 154)
    Else
        chkGuiaAparte.BackColor = &H80000002
    End If
End Sub

Private Function ValidarCampos() As Boolean
    ValidarCampos = False

    If txtFecha.Value = "" Or Not IsDate(txtFecha.Value) Then
        MsgBox "Ingresá una fecha válida en formato DD/MM/AAAA.", vbExclamation
        txtFecha.SetFocus: Exit Function
    End If

    If txtCliente.Value = "" Or txtCodigo.Value = "" Or txtCantidad.Value = "" Then
        MsgBox "Completá todos los campos antes de agregar.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(txtCantidad.Value) Then
        MsgBox "La cantidad debe ser un número.", vbExclamation
        txtCantidad.SetFocus: Exit Function
    End If

    ValidarCampos = True
End Function

Private Sub MostrarItemEnLista(item As Collection)
    lstResumen.AddItem
    With lstResumen
        .List(.ListCount - 1, 0) = item(1)
        .List(.ListCount - 1, 1) = item(2)
        .List(.ListCount - 1, 2) = item(3)
        .List(.ListCount - 1, 3) = item(4)
        .List(.ListCount - 1, 4) = IIf(item(5), "X", "")
        .List(.ListCount - 1, 5) = item(6)
    End With
End Sub

Private Sub btnConfirmar_Click()
    If datosTemp.Count = 0 Then
        MsgBox "No hay ítems para confirmar.", vbExclamation
        Exit Sub
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
    lstResumen.RemoveItem idx
    Call RecalcularPosiciones

    MsgBox "Ítem eliminado y posiciones actualizadas.", vbInformation
    ActualizarEstadoBotones
End Sub

Private Sub btnModificar_Click()
    Dim idx As Long: idx = lstResumen.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná un pedido para modificar.", vbExclamation
        Exit Sub
    End If

    Dim item As Collection: Set item = datosTemp(idx + 1)

    txtFecha.Value = item(1)
    txtCliente.Value = item(2)
    txtCodigo.Value = item(3)
    txtCantidad.Value = item(4)
    chkGuiaAparte.Value = item(5)

    datosTemp.Remove idx + 1
    lstResumen.RemoveItem idx

    Call RecalcularPosiciones
    txtCodigo.SetFocus

    ActualizarEstadoBotones
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub LimpiarTodo()
    Set datosTemp = New Collection
    posicionActual = 10
    txtFecha.Value = ""
    txtCliente.Value = ""
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    chkGuiaAparte.Value = False
    lstResumen.Clear
    txtFecha.SetFocus
    ActualizarEstadoBotones
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
    MsgBox "Nuevo cliente cargado. Puede comenzar otro pedido.", vbInformation
End Sub

Private Sub GuardarYLimpiar()
    Dim hoja As Worksheet: Set hoja = ThisWorkbook.Sheets("ZPVA")
    Dim fila As Long
    fila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row
    If fila < 2 Then fila = 2 Else fila = fila + 1

    Const colPedido As String = "A"
    Const colCliente As String = "B"
    Const colGuia As String = "D"
    Const colFecha As String = "J"
    Const colCodigo As String = "L"
    Const colCantidad As String = "O"
    Const colPosicion As String = "K"

    Dim hayGuiaAparte As Boolean
    Dim item As Collection

    For Each item In datosTemp
        If item(5) = True Then
            hayGuiaAparte = True
            Exit For
        End If
    Next item

    For Each item In datosTemp
        hoja.Cells(fila, colPedido).Value = numeroPedidoActual
        hoja.Cells(fila, colFecha).Value = Format(CDate(item(1)), "yyyymmdd")
        hoja.Cells(fila, colCliente).Value = item(2)
        hoja.Cells(fila, colCodigo).Value = item(3)
        hoja.Cells(fila, colCantidad).Value = item(4)
        hoja.Cells(fila, colPosicion).Value = item(6)
        If hayGuiaAparte Then hoja.Cells(fila, colGuia).Value = "X"
        fila = fila + 1
    Next item

    numeroPedidoActual = numeroPedidoActual + 1
    Call LimpiarTodo
End Sub

Private Sub txtFecha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCliente.SetFocus
End Sub

Private Sub txtCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCodigo.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCantidad.SetFocus
End Sub

Private Sub txtCantidad_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnAgregarOtro.SetFocus
End Sub

Private Sub txtFecha_Change()
    Dim txt As String
    Static editando As Boolean
    If editando Then Exit Sub

    txt = Replace(txtFecha.Text, "/", "")
    If Not IsNumeric(txt) Then
        txtFecha.Text = ""
        Exit Sub
    End If

    editando = True
    Select Case Len(txt)
        Case 1 To 2
            txtFecha.Text = txt
        Case 3 To 4
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3)
        Case 5 To 8
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2) & "/" & Mid(txt, 5)
        Case Is > 8
            txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2) & "/" & Mid(txt, 5, 4)
    End Select
    txtFecha.SelStart = Len(txtFecha.Text)
    editando = False
End Sub

Private Sub lstResumen_Change()
    ActualizarEstadoBotones
End Sub

Private Sub ActualizarEstadoBotones()
    Dim tieneSeleccion As Boolean
    tieneSeleccion = (lstResumen.ListIndex <> -1)
    btnModificar.Enabled = tieneSeleccion
    btnEliminar.Enabled = tieneSeleccion
End Sub

