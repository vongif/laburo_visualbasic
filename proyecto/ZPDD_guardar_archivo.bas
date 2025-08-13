Attribute VB_Name = "ZPDD_guardar_archivo"

Sub GuardarHojaArchivoCompletaComoTXT()
    Dim hoja As Worksheet
    Dim rutaArchivo As String
    Dim archivo As Integer
    Dim fila As Long, col As Long
    Dim ultimaFila As Long, ultimaCol As Long
    Dim linea As String
    Dim dlg As FileDialog
    Dim codigo As String, cantidad As String
    Dim valorCelda As String

    Set hoja = ThisWorkbook.Sheets("ZPDD_devo_minorista")

    ' Crear diálogo para seleccionar ruta y nombre del archivo
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    With dlg
        .Title = "Guardar hoja 'ZPDD_devo_minorista' como TXT"
        .InitialFileName = "archivo_completo_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
        If .Show <> -1 Then
            MsgBox "Operación cancelada.", vbExclamation
            Exit Sub
        End If
        rutaArchivo = .SelectedItems(1)
    End With

    ' Asegurar que tenga extensión .txt (reemplaza si tiene otra extensión)
    If InStrRev(rutaArchivo, ".") > 0 Then
        rutaArchivo = Left(rutaArchivo, InStrRev(rutaArchivo, ".") - 1)
    End If
    rutaArchivo = rutaArchivo & ".txt"

    ' Detectar último rango usado
    ultimaFila = hoja.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ultimaCol = hoja.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Crear archivo de texto
    archivo = FreeFile
    Open rutaArchivo For Output As #archivo

    ' Escribir encabezado (fila 1)
    linea = ""
    For col = 1 To ultimaCol
        linea = linea & Trim(hoja.Cells(1, col).Text) & vbTab
    Next col
    If Right(linea, 1) = vbTab Then linea = Left(linea, Len(linea) - 1)
    Print #archivo, linea

    ' Guardar filas con datos en columnas L (12) y O (15)
    For fila = 2 To ultimaFila
        codigo = Trim(hoja.Cells(fila, 12).Value)   ' Columna L
        cantidad = Trim(hoja.Cells(fila, 15).Value) ' Columna O

        If codigo <> "" And cantidad <> "" Then
            linea = ""
            For col = 1 To ultimaCol
                If col = 16 Then ' Columna P - Fecha Entrega
                    If IsDate(hoja.Cells(fila, col).Value) Then
                        valorCelda = Format(hoja.Cells(fila, col).Value, "yyyymmdd")
                    Else
                        valorCelda = Trim(hoja.Cells(fila, col).Text)
                    End If
                Else
                    valorCelda = Trim(hoja.Cells(fila, col).Text)
                End If
                linea = linea & valorCelda & vbTab
            Next col
            If Right(linea, 1) = vbTab Then linea = Left(linea, Len(linea) - 1)
            Print #archivo, linea
        End If
    Next fila

    Close #archivo

    MsgBox "Se guardaron las filas válidas con encabezado en:" & vbCrLf & rutaArchivo, vbInformation
End Sub


