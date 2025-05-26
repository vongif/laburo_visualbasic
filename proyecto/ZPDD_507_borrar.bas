Attribute VB_Name = "ZPDD_507_borrar"
Sub BorrarDatosIngresadosSeleccionados()
    Dim hoja As Worksheet
    Dim ultimaFila As Long
    Set hoja = ThisWorkbook.Sheets("ZPDD_507")
    
    If MsgBox("¿Querés borrar los datos ingresados?", vbYesNo + vbQuestion, "Confirmar borrado") = vbNo Then
        Exit Sub
    End If
    
    ' Determinar la última fila con datos en las columnas clave
    ultimaFila = Application.WorksheetFunction.Max( _
        hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "E").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "K").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "L").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "O").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "D").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "H").End(xlUp).Row _
    )
    
    If ultimaFila < 2 Then
        MsgBox "No hay datos para borrar.", vbInformation
        Exit Sub
    End If

    ' Borrar contenido solo hasta la última fila usada
    hoja.Range("A2:A" & ultimaFila).ClearContents   ' N° Pedido
    hoja.Range("B2:B" & ultimaFila).ClearContents   ' Cliente
    hoja.Range("D2:D" & ultimaFila).ClearContents   ' Guía aparte
    hoja.Range("E2:E" & ultimaFila).ClearContents   ' Marcas
    hoja.Range("H2:H" & ultimaFila).ClearContents   ' Remito ? NUEVO
    hoja.Range("K2:K" & ultimaFila).ClearContents   ' Posición
    hoja.Range("L2:L" & ultimaFila).ClearContents   ' Código
    hoja.Range("O2:O" & ultimaFila).ClearContents   ' Cantidad

    MsgBox "Datos borrados correctamente.", vbInformation
End Sub


