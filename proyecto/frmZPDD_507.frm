VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZPDD_507 
   Caption         =   "ZPDD_507"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11070
   OleObjectBlob   =   "frmZPDD_507.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZPDD_507"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================
' VARIABLES GLOBALES
' ===========================
Dim numeroPedidoActual As Long
Dim posicionActual As Long
Dim datosTemp As Collection
Dim marcasPedidoActual As String ' NUEVO

' ===========================
' INICIALIZACIÓN
' ===========================
Private Sub UserForm_Initialize()
    Dim hoja As Worksheet
    Set hoja = ThisWorkbook.Sheets("ZPDD_507")

    Dim ultimaFila As Long
    ultimaFila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 2 Then
        numeroPedidoActual = 1
    ElseIf IsNumeric(hoja.Cells(ultimaFila, 1).Value) Then
        numeroPedidoActual = hoja.Cells(ultimaFila, 1).Value + 1
    Else
        numeroPedidoActual = 1
    End If

    posicionActual = 10
    Set datosTemp = New Collection
    marcasPedidoActual = ""

    With lstResumen
        .ColumnCount = 6
        .ColumnWidths = "60 pt;90pt;90 pt;90 pt;80 pt;80 pt"
        
    End With
        
    txtFecha.SetFocus
    btnModificar.Enabled = False
    btnEliminar.Enabled = False
    
    
End Sub

' ===========================
' EVENTOS DE BOTONES
' ===========================
Private Sub btnAgregarOtro_Click()
    If Not ValidarCampos() Then Exit Sub
    
    ' Validación de cantidad que no empiece con cero
    If Left(txtCantidad.Text, 1) = "0" And Len(txtCantidad.Text) > 1 Then
        MsgBox "La cantidad no puede comenzar con cero.", vbExclamation
        txtCantidad.SetFocus
        Exit Sub
    End If

    Dim item As New Collection
    item.Add Format(CDate(txtFecha.Value), "dd/mm/yyyy")
    item.Add txtCliente.Value
    item.Add txtRemito.Value
    item.Add txtCodigo.Value
    item.Add txtCantidad.Value
    item.Add IIf(chkGuiaAparte.Value, "X", "")
    item.Add ObtenerMarcasSeleccionadas()
    item.Add posicionActual

    datosTemp.Add item
    MostrarItemEnLista item

    posicionActual = posicionActual + 10
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    txtCodigo.SetFocus
End Sub

Private Sub btnConfirmar_Click()
    If datosTemp.Count = 0 Then
        MsgBox "No hay ítems para confirmar.", vbExclamation
        Exit Sub
    End If
    GuardarYLimpiar
    MsgBox "Pedido confirmado y cargado en la planilla.", vbInformation
End Sub

Private Sub btnEliminar_Click()
    Dim idx As Long: idx = lstResumen.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná un pedido para eliminar.", vbExclamation
        Exit Sub
    End If

    datosTemp.Remove idx + 1
    RecalcularPosiciones
    MsgBox "Ítem eliminado y posiciones actualizadas.", vbInformation
End Sub

Private Sub btnModificar_Click()
     Dim idx As Long: idx = lstResumen.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná un pedido para modificar.", vbExclamation
        Exit Sub
    End If

    If Not ValidarCampos() Then Exit Sub

    Dim item As New Collection
    item.Add Format(CDate(txtFecha.Value), "dd/mm/yyyy") ' 1 - Fecha
    item.Add txtCliente.Value                             ' 2 - Cliente
    item.Add txtRemito.Value                              ' 3 - Remito
    item.Add txtCodigo.Value                              ' 4 - Código
    item.Add txtCantidad.Value                            ' 5 - Cantidad
    item.Add IIf(chkGuiaAparte.Value, "X", "")            ' 6 - Guía
    item.Add ObtenerMarcasSeleccionadas()                 ' 7 - Marcas
    item.Add datosTemp(idx + 1)(8)                        ' 8 - Posición (no cambia)

    ' Reemplazar en colección
    Set datosTemp(idx + 1) = item

    ' Reemplazar en ListBox
    With lstResumen
        .List(idx, 0) = item(1) ' Fecha
        .List(idx, 1) = item(2) ' Cliente
        .List(idx, 2) = item(3) ' Remito
        .List(idx, 3) = item(4) ' Código
        .List(idx, 4) = item(5) ' Cantidad
        .List(idx, 5) = item(7) ' Marcas
    End With

    MsgBox "Ítem modificado correctamente.", vbInformation

    ' Limpiar campos después de modificar
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    txtCodigo.SetFocus
End Sub

Private Sub lstResumen_Change()
    btnModificar.Enabled = (lstResumen.ListIndex <> -1)
    btnEliminar.Enabled = (lstResumen.ListIndex <> -1)
End Sub

Private Sub btnNuevoCliente_Click()
    If datosTemp.Count = 0 Then
        MsgBox "No hay ítems para guardar.", vbExclamation
        Exit Sub
    End If
    GuardarYLimpiar
    numeroPedidoActual = numeroPedidoActual + 1
    MsgBox "Nuevo cliente iniciado. Número de pedido: " & numeroPedidoActual, vbInformation
    
End Sub

Private Sub btnNuevoPedido_Click()
    numeroPedidoActual = numeroPedidoActual + 1
    LimpiarTodo
    MsgBox "Nuevo pedido iniciado. Número de pedido: " & numeroPedidoActual, vbInformation
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

' ===========================
' VALIDACIONES Y UTILITARIOS
' ===========================
Private Function ValidarCampos() As Boolean
    ValidarCampos = False

    If txtFecha.Value = "" Or Not IsDate(txtFecha.Value) Then
        MsgBox "Ingresá una fecha válida en formato DD/MM/AAAA.", vbExclamation
        txtFecha.SetFocus: Exit Function
    End If

    If txtCliente.Value = "" Or txtCodigo.Value = "" Or txtCantidad.Value = "" Or txtRemito.Value = "" Then
        MsgBox "Completá todos los campos antes de agregar.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(txtCantidad.Value) Then
        MsgBox "La cantidad debe ser un número.", vbExclamation
        txtCantidad.SetFocus: Exit Function
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
         .List(.ListCount - 1, 0) = item(1) ' Fecha
        .List(.ListCount - 1, 1) = item(2) ' Cliente
        .List(.ListCount - 1, 2) = item(3) ' Remito
        .List(.ListCount - 1, 3) = item(4) ' Material
        .List(.ListCount - 1, 4) = item(5) ' Cantidad
        .List(.ListCount - 1, 5) = item(7) ' Marcas / Org
        ' No mostrar item(8) (posición)
           
      End With
End Sub

Private Sub RecalcularPosiciones()
     Dim i As Long
    Dim nuevaPosicion As Long: nuevaPosicion = 10
    Dim item As Collection
    Dim nuevoItem As Collection
    Dim nuevaLista As New Collection

    lstResumen.Clear

    For i = 1 To datosTemp.Count
        Set item = datosTemp(i)
        Set nuevoItem = New Collection

        nuevoItem.Add item(1) ' Fecha
        nuevoItem.Add item(2) ' Cliente
        nuevoItem.Add item(3) ' Remito
        nuevoItem.Add item(4) ' Código
        nuevoItem.Add item(5) ' Cantidad
        nuevoItem.Add item(6) ' Guía aparte
        nuevoItem.Add item(7) ' Marcas
        nuevoItem.Add nuevaPosicion ' Nueva posición

        nuevaLista.Add nuevoItem
        MostrarItemEnLista nuevoItem
        nuevaPosicion = nuevaPosicion + 10
    Next i

    ' Reemplazamos la colección original por la nueva con las posiciones corregidas
    Set datosTemp = nuevaLista
    posicionActual = nuevaPosicion
End Sub

Private Sub GuardarYLimpiar()
    Dim hoja As Worksheet: Set hoja = ThisWorkbook.Sheets("ZPDD_507")
    Dim fila As Long
    fila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row
    If fila < 2 Then fila = 2 Else fila = fila + 1

    Const colPedido As Long = 1
    Const colCliente As Long = 2
    Const colGuia As Long = 4
    Const colFecha As Long = 10
    Const colCodigo As Long = 12
    Const colPosicion As Long = 11
    Const colCantidad As Long = 15
    Const colMarcas As Long = 5
    Const colRemito As Long = 8

    Dim item As Collection
    For Each item In datosTemp
        hoja.Cells(fila, colPedido).Value = numeroPedidoActual
        hoja.Cells(fila, colFecha).Value = Format(CDate(item(1)), "yyyymmdd")
        hoja.Cells(fila, colCliente).Value = item(2)
        hoja.Cells(fila, colCodigo).Value = item(4)
        hoja.Cells(fila, colCantidad).Value = item(5)
        hoja.Cells(fila, colPosicion).Value = item(8)
        hoja.Cells(fila, colMarcas).Value = item(7)
        hoja.Cells(fila, colRemito).Value = item(3)
        fila = fila + 1
    Next item

    hoja.Range("Z1").Value = numeroPedidoActual
    LimpiarTodo
End Sub

Private Sub LimpiarTodo()
    Set datosTemp = New Collection
    posicionActual = 10
    txtFecha.Value = ""
    txtCliente.Value = ""
    txtCodigo.Value = ""
    txtCantidad.Value = ""
    txtRemito.Value = ""
    chkGuiaAparte.Value = False
    lstResumen.Clear
    marcasPedidoActual = ""
    txtFecha.SetFocus
    LimpiarChecks
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




' ===========================
' NAVEGACIÓN CON ENTER
' ===========================
Private Sub txtFecha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCliente.SetFocus
End Sub

Private Sub txtCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtRemito.SetFocus
End Sub

Private Sub txtRemito_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
        Case 1 To 2: txtFecha.Text = txt
        Case 3 To 4: txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3)
        Case 5 To 8: txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2) & "/" & Mid(txt, 5)
        Case Is > 8: txtFecha.Text = Left(txt, 2) & "/" & Mid(txt, 3, 2) & "/" & Mid(txt, 5, 4)
    End Select
    txtFecha.SelStart = Len(txtFecha.Text)
    editando = False
End Sub


