Attribute VB_Name = "Compra_ESP"

Sub Compra()
'Queremos saber el porcentaje de en qué tiendas compran los clientes de cada una de las marcas

    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long, ultimaColumna As Long
    Dim columnaCompra As Long
    Dim filaPregunta As Long
    Dim marcas As Variant
    Dim colDestino As Long, filaDestino As Long
    Dim i As Long, j As Long, k As Long
    Dim filaOrigen As Long, columnaOrigen As Long
    Dim filaBases As Long
    Dim colDestino_bases As Long


    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("Tablas")
    Set wsDestino = ThisWorkbook.Sheets("Compra")

    ' Encuentra la última fila y columna en la hoja origen
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaColumna = wsOrigen.Cells(7, wsOrigen.Columns.Count).End(xlToLeft).Column

    ' Encuentra la columna donde aparece "Compra". Aqui vamos a tener una columna para cada marca (clientes de esa marca)
    columnaCompra = 0
    For j = 1 To ultimaColumna
        If Trim(wsOrigen.Cells(7, j).Value) = "Compra" Then
            columnaCompra = j
            Exit For
        End If
    Next j

'Si no lo encuentra, devuelve un mensaje indicando que no se ha encontrado
    If columnaCompra = 0 Then
        MsgBox "No se encontró la columna con 'Compra' en la fila 7.", vbExclamation
        Exit Sub
    End If

    ' Define las marcas
    marcas = Array("Nike", "Adidas", "Vans", "Reebok", "Asics", _
                   "Puma", "Levi's", "Decathlon")

    ' Borra el contenido del rango de destino
    wsDestino.Range("B2:N10").ClearContents
    wsDestino.Range("C14:N14").ClearContents

    ' Escribe las marcas como encabezados de filas y columnas en la hoja destino
    colDestino = 3 ' Inicia en la columna 3
    filaDestino = 3 ' Inicia en la fila 3

    For i = LBound(marcas) To UBound(marcas)
        wsDestino.Cells(2, colDestino + i).Value = marcas(i) ' Encabezados de columna
        wsDestino.Cells(filaDestino + i, 2).Value = marcas(i) ' Encabezados de fila
    Next i

    ' Encuentra la fila donde aparece "Pregunta - COMPRA" en la columna A
    filaPregunta = 0
    For i = 1 To ultimaFila
        If Trim(wsOrigen.Cells(i, 1).Value) = "Pregunta - COMPRA" Then
            filaPregunta = i
            Exit For
        End If
    Next i

'Si no lo encuentra
    If filaPregunta = 0 Then
        MsgBox "No se encontró la pregunta 'Pregunta - COMPRA.", vbExclamation
        Exit Sub
    End If

    ' Llena la matriz de la hoja destino
    For i = LBound(marcas) To UBound(marcas)
        For j = LBound(marcas) To UBound(marcas)
            filaOrigen = 0
            columnaOrigen = 0

            ' Busca la fila de la marca origen
            For k = filaPregunta + 1 To ultimaFila
                If Trim(wsOrigen.Cells(k, 1).Value) = marcas(i) Then
                    filaOrigen = k
                    Exit For
                End If
            Next k

            ' Busca la columna de la marca destino
            For k = columnaCompra To ultimaColumna
                If Trim(wsOrigen.Cells(9, k).Value) = marcas(j) Then
                    columnaOrigen = k
                    Exit For
                End If
            Next k

            ' Llena la celda correspondiente
            If filaOrigen > 0 And columnaOrigen > 0 Then
                wsDestino.Cells(filaDestino + i, colDestino + j).Value = wsOrigen.Cells(filaOrigen, columnaOrigen).Value
            Else
                wsDestino.Cells(filaDestino + i, colDestino + j).Value = 0
            End If
        Next j
    Next i
End Sub



