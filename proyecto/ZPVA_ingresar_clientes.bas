Attribute VB_Name = "ZPVA_ingresar_clientes"
Sub IngresarMultiplesClientes()
    Dim hoja As Worksheet
    Dim fila As Long
    Dim codigo As String
    Dim cantidad As Variant
    Dim cliente As String
    Dim continuar As VbMsgBoxResult
    Dim i As Long
    Dim codigoYaExiste As Boolean

    Set hoja = ThisWorkbook.Sheets("ZPVA")

    ' Buscar la próxima fila vacía en columna L
    fila = hoja.Cells(hoja.Rows.Count, "L").End(xlUp).Row
    If fila < 2 Then
        fila = 2
    Else
        fila = fila + 1
    End If

    Do
        ' Pedir número de cliente
        cliente = InputBox("Ingresá el número de CLIENTE:", "Cliente")
        If cliente = "" Then
            MsgBox "Carga cancelada.", vbExclamation
            Exit Sub
        End If

        ' Cargar códigos para este cliente
        Do
            codigo = InputBox("Ingresá el CÓDIGO del material (dejar vacío para pasar al siguiente cliente):", "Código para cliente " & cliente)
            If codigo = "" Then Exit Do

            ' Verificar si el código ya existe para este cliente
            codigoYaExiste = False
            For i = 2 To hoja.Cells(hoja.Rows.Count, "L").End(xlUp).Row
                If hoja.Cells(i, "B").Value = cliente And hoja.Cells(i, "L").Value = codigo Then
                    codigoYaExiste = True
                    Exit For
                End If
            Next i

            If codigoYaExiste Then
                MsgBox "El código '" & codigo & "' ya fue ingresado para el cliente " & cliente & ".", vbExclamation
                GoTo ContinuarCarga
            End If

            cantidad = InputBox("Ingresá la CANTIDAD correspondiente:", "Cantidad para " & codigo)
            If cantidad = "" Then Exit Do

            If Not IsNumeric(cantidad) Then
                MsgBox "La cantidad debe ser un número.", vbExclamation
                GoTo ContinuarCarga
            End If

            hoja.Cells(fila, "B").Value = cliente      ' Cliente
            hoja.Cells(fila, "L").Value = codigo       ' Código
            hoja.Cells(fila, "O").Value = cantidad     ' Cantidad

            fila = fila + 1

ContinuarCarga:
        Loop

        continuar = MsgBox("¿Querés cargar otro cliente?", vbYesNo + vbQuestion, "Continuar")
        If continuar = vbNo Then Exit Do
    Loop

    MsgBox "Carga finalizada correctamente.", vbInformation
End Sub

