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

    ' Buscar la pr�xima fila vac�a en columna L
    fila = hoja.Cells(hoja.Rows.Count, "L").End(xlUp).Row
    If fila < 2 Then
        fila = 2
    Else
        fila = fila + 1
    End If

    Do
        ' Pedir n�mero de cliente
        cliente = InputBox("Ingres� el n�mero de CLIENTE:", "Cliente")
        If cliente = "" Then
            MsgBox "Carga cancelada.", vbExclamation
            Exit Sub
        End If

        ' Cargar c�digos para este cliente
        Do
            codigo = InputBox("Ingres� el C�DIGO del material (dejar vac�o para pasar al siguiente cliente):", "C�digo para cliente " & cliente)
            If codigo = "" Then Exit Do

            ' Verificar si el c�digo ya existe para este cliente
            codigoYaExiste = False
            For i = 2 To hoja.Cells(hoja.Rows.Count, "L").End(xlUp).Row
                If hoja.Cells(i, "B").Value = cliente And hoja.Cells(i, "L").Value = codigo Then
                    codigoYaExiste = True
                    Exit For
                End If
            Next i

            If codigoYaExiste Then
                MsgBox "El c�digo '" & codigo & "' ya fue ingresado para el cliente " & cliente & ".", vbExclamation
                GoTo ContinuarCarga
            End If

            cantidad = InputBox("Ingres� la CANTIDAD correspondiente:", "Cantidad para " & codigo)
            If cantidad = "" Then Exit Do

            If Not IsNumeric(cantidad) Then
                MsgBox "La cantidad debe ser un n�mero.", vbExclamation
                GoTo ContinuarCarga
            End If

            hoja.Cells(fila, "B").Value = cliente      ' Cliente
            hoja.Cells(fila, "L").Value = codigo       ' C�digo
            hoja.Cells(fila, "O").Value = cantidad     ' Cantidad

            fila = fila + 1

ContinuarCarga:
        Loop

        continuar = MsgBox("�Quer�s cargar otro cliente?", vbYesNo + vbQuestion, "Continuar")
        If continuar = vbNo Then Exit Do
    Loop

    MsgBox "Carga finalizada correctamente.", vbInformation
End Sub

