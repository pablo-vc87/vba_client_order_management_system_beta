Attribute VB_Name = "ModuloWord"
Option Explicit

' Genera la orden de compra y la orden de trabajo en Word a partir de un pedido.
' Las plantillas llevan marcadores {{CAMPO}} y UN bloque modelo de productos: la fila
' con {{TECNICA}} y, opcionalmente, filas debajo con {{OBS_PRODUCTO}} o {{RENDER}}.
' La macro repite ese bloque una vez por producto, así que no hay límite de productos.
' Usa EjecutarSQL (ModuloSQL) y RutaAbsolutaImagen (ModuloImagenes).

Private Const PLANTILLA_OC As String = "Plantilla_Orden_Compra.docx"
Private Const PLANTILLA_OT As String = "Plantilla_Orden_Trabajo.docx"
Private Const CARPETA_ORDENES As String = "ordenes"
Private Const ANCHO_RENDER As Single = 90       'imagen dentro de una columna (~3.2 cm)
Private Const ANCHO_MAX_RENDER As Single = 360  'imagen en una fila completa (~12.7 cm)
Private Const ALTO_MAX_RENDER As Single = 200   'alto máximo en una fila completa (~7 cm)
Private Const MAX_CELDAS As Long = 30

Private mObservaciones As String                 'observaciones de todos los productos

Public Sub GenerarOrdenCompra(ByVal PedidoID As Long)
    GenerarOrden PedidoID, PLANTILLA_OC, "OC"
End Sub

Public Sub GenerarOrdenTrabajo(ByVal PedidoID As Long)
    GenerarOrden PedidoID, PLANTILLA_OT, "OT"
End Sub

'Banco de pruebas: pide el número de pedido y genera las dos órdenes
Public Sub ProbarOrdenes()
    Dim r As String
    r = InputBox("Número de pedido:", "CALEIDO")
    If Not IsNumeric(r) Then Exit Sub
    GenerarOrdenCompra CLng(r)
    GenerarOrdenTrabajo CLng(r)
End Sub

' ===============================================
' Llena una plantilla con el pedido y la guarda en .docx y .pdf
' ===============================================
Private Sub GenerarOrden(ByVal PedidoID As Long, ByVal plantilla As String, ByVal prefijo As String)
    Dim rsPed As Object, rsDet As Object
    Dim wd As Object, doc As Object, tbl As Object, bloque As Object, destino As Object
    Dim rutaPlantilla As String, carpeta As String, archivo As String
    Dim filaModelo As Long, nFilas As Long, n As Long, i As Long, k As Long, c As Long
    Dim Fila As Long, nCeldas As Long, conContenido As Boolean
    Dim modelo() As String, obs As String

    rutaPlantilla = ThisWorkbook.Path & "\" & plantilla
    If Dir(rutaPlantilla) = "" Then
        MsgBox "No se encontró la plantilla:" & vbCrLf & rutaPlantilla, vbExclamation, "CALEIDO"
        Exit Sub
    End If

    '1. Datos del pedido con SQL (ACE lee el archivo del disco: guardar antes)
    ThisWorkbook.Save
    Set rsPed = EjecutarSQL("SELECT * FROM [Pedidos$] WHERE PedidoID = " & PedidoID)
    If rsPed Is Nothing Then Exit Sub
    If rsPed.EOF Then
        MsgBox "No existe el pedido " & PedidoID, vbExclamation, "CALEIDO"
        Exit Sub
    End If
    Set rsDet = EjecutarSQL("SELECT * FROM [Detalle_Pedidos$] WHERE PedidoID = " & PedidoID & _
                            " ORDER BY NoProducto")
    If rsDet Is Nothing Then Exit Sub
    n = rsDet.RecordCount

    '2. Abrir Word y crear un documento nuevo a partir de la plantilla
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then Set wd = CreateObject("Word.Application")

    On Error GoTo Fallo
    Set doc = wd.Documents.Add(Template:=rutaPlantilla)

    If Not BuscarFilaModelo(doc, tbl, filaModelo) Then
        MsgBox "La plantilla no tiene una fila con {{TECNICA}}.", vbExclamation, "CALEIDO"
        doc.Close False
        Exit Sub
    End If

    '3. Guardar el texto del bloque modelo y repetirlo (un bloque por producto)
    nFilas = TamanoBloque(tbl, filaModelo)
    ReDim modelo(1 To nFilas, 1 To MAX_CELDAS)
    For k = 1 To nFilas
        For c = 1 To tbl.Rows(filaModelo + k - 1).Cells.Count
            modelo(k, c) = TextoCelda(tbl.Rows(filaModelo + k - 1).Cells(c))
        Next c
    Next k

    If n = 0 Then
        For k = nFilas To 1 Step -1
            tbl.Rows(filaModelo + k - 1).Delete
        Next k
    Else
        'Cada copia se inserta arriba del bloque modelo (FormattedText copia las filas
        'con su formato y sus celdas combinadas)
        For i = 2 To n
            Set bloque = doc.Range(tbl.Rows(filaModelo).Range.Start, _
                                   tbl.Rows(filaModelo + nFilas - 1).Range.End)
            Set destino = doc.Range(bloque.Start, bloque.Start)
            destino.FormattedText = bloque.FormattedText
        Next i
    End If

    '4. Llenar cada bloque con su producto; las filas extra sin contenido se borran
    Fila = filaModelo
    Do While Not rsDet.EOF
        For k = 1 To nFilas
            nCeldas = tbl.Rows(Fila).Cells.Count
            conContenido = (k = 1)                  'la fila principal siempre se queda
            For c = 1 To nCeldas
                If LlenarCelda(tbl.Rows(Fila).Cells(c), modelo(k, c), rsPed, rsDet, _
                               nCeldas = 1) Then conContenido = True
            Next c
            If conContenido Then
                Fila = Fila + 1
            Else
                tbl.Rows(Fila).Delete
            End If
        Next k
        If Texto(Campo(rsDet, "Observaciones")) <> "" Then
            If obs <> "" Then obs = obs & "; "
            obs = obs & "Producto " & Texto(Campo(rsDet, "NoProducto")) & ": " & _
                  Texto(Campo(rsDet, "Observaciones"))
        End If
        rsDet.MoveNext
    Loop
    mObservaciones = obs

    '5. Datos del encabezado (lo que queda fuera de la tabla de productos)
    ReemplazarEnDocumento doc, rsPed

    '6. Guardar en la carpeta "ordenes" junto al libro, sin sobrescribir versiones anteriores
    carpeta = ThisWorkbook.Path & "\" & CARPETA_ORDENES
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta
    archivo = NombreLibre(carpeta & "\" & prefijo & "_" & Format(PedidoID, "0000") & "_" & _
              NombreArchivo(Texto(Campo(rsPed, "RazonSocial"))))
    doc.SaveAs2 archivo & ".docx", 16          'wdFormatDocumentDefault
    doc.ExportAsFixedFormat archivo & ".pdf", 17   'wdExportFormatPDF

    wd.Visible = True
    doc.Activate
    Exit Sub

Fallo:
    MsgBox "No se pudo generar la orden:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Si el archivo ya estaba abierto en Word, ciérralo e inténtalo de nuevo.", _
           vbCritical, "CALEIDO"
    If Not wd Is Nothing Then wd.Visible = True
End Sub

' ===============================================
' Busca la tabla y la fila que contienen {{TECNICA}}
' ===============================================
Private Function BuscarFilaModelo(ByVal doc As Object, ByRef tbl As Object, _
                                  ByRef Fila As Long) As Boolean
    Dim t As Object, r As Long
    For Each t In doc.Tables
        For r = 1 To t.Rows.Count
            If InStr(t.Rows(r).Range.Text, "{{TECNICA}}") > 0 Then
                Set tbl = t
                Fila = r
                BuscarFilaModelo = True
                Exit Function
            End If
        Next r
    Next t
End Function

' Filas del bloque modelo: la de {{TECNICA}} más las que siguen con {{OBS_PRODUCTO}} o {{RENDER}}
Private Function TamanoBloque(ByVal tbl As Object, ByVal filaModelo As Long) As Long
    Dim r As Long, t As String
    TamanoBloque = 1
    For r = filaModelo + 1 To tbl.Rows.Count
        t = tbl.Rows(r).Range.Text
        If InStr(t, "{{OBS_PRODUCTO}}") = 0 And InStr(t, "{{RENDER}}") = 0 Then Exit For
        TamanoBloque = TamanoBloque + 1
    Next r
End Function

' Texto de una celda sin la marca de fin de celda (Chr(13) & Chr(7))
Private Function TextoCelda(ByVal cel As Object) As String
    Dim t As String
    t = cel.Range.Text
    TextoCelda = Left(t, Len(t) - 2)
End Function

' Escribe en la celda el texto del modelo con los marcadores reemplazados; en la celda
' {{RENDER}} inserta la imagen del producto (grande si la celda ocupa toda la fila).
' Devuelve True si algún marcador tuvo valor o se insertó una imagen.
Private Function LlenarCelda(ByVal cel As Object, ByVal modelo As String, _
                             ByVal rsPed As Object, ByVal rsDet As Object, _
                             ByVal filaCompleta As Boolean) As Boolean
    Dim ruta As String, img As Object, conValor As Boolean

    If InStr(modelo, "{{RENDER}}") > 0 Then
        cel.Range.Text = ""
        ruta = RutaAbsolutaImagen(Texto(Campo(rsDet, "RutaRender")))
        If ruta <> "" Then
            If Dir(ruta) <> "" Then
                Set img = cel.Range.InlineShapes.AddPicture(Filename:=ruta, _
                          LinkToFile:=False, SaveWithDocument:=True)
                img.LockAspectRatio = True
                If filaCompleta Then
                    AjustarImagen img, ANCHO_MAX_RENDER, ALTO_MAX_RENDER
                Else
                    img.Width = ANCHO_RENDER
                End If
                LlenarCelda = True
            End If
        End If
    Else
        cel.Range.Text = ReemplazarMarcadores(modelo, rsPed, rsDet, conValor)
        LlenarCelda = conValor
    End If
End Function

' Reduce la imagen (sin deformarla) para que quepa en anchoMax x altoMax
Private Sub AjustarImagen(ByVal img As Object, ByVal anchoMax As Single, ByVal altoMax As Single)
    Dim w As Single, h As Single, f As Single
    w = img.Width
    h = img.Height
    If w <= 0 Or h <= 0 Then Exit Sub
    f = anchoMax / w
    If altoMax / h < f Then f = altoMax / h
    If f < 1 Then
        img.Width = w * f
        img.Height = h * f
    End If
End Sub

' Reemplaza cada {{CAMPO}} de un texto por su valor; conValor indica si alguno tuvo valor
Private Function ReemplazarMarcadores(ByVal texto As String, ByVal rsPed As Object, _
                                      ByVal rsDet As Object, ByRef conValor As Boolean) As String
    Dim p As Long, q As Long, valorCampo As String

    p = InStr(texto, "{{")
    Do While p > 0
        q = InStr(p, texto, "}}")
        If q = 0 Then Exit Do
        valorCampo = Valor(Mid(texto, p + 2, q - p - 2), rsPed, rsDet)
        If valorCampo <> "" Then conValor = True
        texto = Left(texto, p - 1) & valorCampo & Mid(texto, q + 2)
        p = InStr(p + Len(valorCampo), texto, "{{")
    Loop
    ReemplazarMarcadores = texto
End Function

' Reemplaza los {{CAMPO}} del cuerpo del documento (encabezado del pedido)
Private Sub ReemplazarEnDocumento(ByVal doc As Object, ByVal rsPed As Object)
    Dim rng As Object, t As String

    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "\{\{[A-Z_]@\}\}"
        .MatchWildcards = True
        .Forward = True
        .Wrap = 0                                  'wdFindStop
    End With
    Do While rng.Find.Execute
        t = rng.Text
        rng.Text = Valor(Mid(t, 3, Len(t) - 4), rsPed, Nothing)
        rng.Collapse 0                             'wdCollapseEnd
    Loop
End Sub

' ===============================================
' Valor de cada marcador. Para agregar uno nuevo a la plantilla, agrégalo aquí.
' ===============================================
Private Function Valor(ByVal nombre As String, ByVal rsPed As Object, _
                       ByVal rsDet As Object) As String
    Select Case nombre
        'Pedido
        Case "PEDIDO":            Valor = Texto(Campo(rsPed, "PedidoID"))
        Case "FECHA":             Valor = Fecha(Campo(rsPed, "FechaPedido"))
        Case "CONTACTO":          Valor = Texto(Campo(rsPed, "NombreCliente"))
        Case "RAZON_SOCIAL":      Valor = Texto(Campo(rsPed, "RazonSocial"))
        Case "DOMICILIO_FISCAL":  Valor = Texto(Campo(rsPed, "DomicilioFiscal"))
        Case "EMAIL":             Valor = Texto(Campo(rsPed, "Email"))
        Case "TELEFONO":          Valor = Unir(Texto(Campo(rsPed, "Telefono")), _
                                               Texto(Campo(rsPed, "Telefono2")), " / ")
        Case "FECHA_ENTREGA":     Valor = Fecha(Campo(rsPed, "FechaEntrega"))
        Case "DIRECCION_ENTREGA": Valor = Texto(Campo(rsPed, "DireccionEntrega"))
        Case "TOTAL_PEDIDO":      Valor = Dinero(Campo(rsPed, "TotalPedido"))
        Case "OBSERVACIONES":     Valor = mObservaciones
        Case Else
            'Producto
            If rsDet Is Nothing Then Exit Function
            Select Case nombre
                Case "TECNICA":         Valor = Texto(Campo(rsDet, "Tecnica"))
                Case "MATERIAL":        Valor = Texto(Campo(rsDet, "Material"))
                Case "CANTIDAD":        Valor = Texto(Campo(rsDet, "Cantidad"))
                Case "PRECIO_UNITARIO": Valor = Dinero(Campo(rsDet, "PrecioUnitario"))
                Case "SUBTOTAL":        Valor = Dinero(Campo(rsDet, "SubtotalProducto"))
                Case "IVA":             Valor = Dinero(Campo(rsDet, "IVAProducto"))
                Case "TOTAL":           Valor = Dinero(Campo(rsDet, "TotalProducto"))
                Case "FECHA_RECEPCION": Valor = Fecha(Campo(rsDet, "FechaRecepcion"))
                Case "PANTONE":         Valor = Texto(Campo(rsDet, "Pantone"))
                Case "LOGO":            Valor = Texto(Campo(rsDet, "NombreLogo"))
                Case "TAMANO_LOGO":     Valor = Texto(Campo(rsDet, "TamanoLogo"))
                Case "OBS_PRODUCTO":    Valor = Texto(Campo(rsDet, "Observaciones"))
            End Select
    End Select
End Function

' ===============================================
' Utilidades de formato
' ===============================================
Private Function Campo(ByVal rs As Object, ByVal nombre As String) As Variant
    Campo = rs.Fields(nombre).Value
End Function

Private Function Texto(ByVal v As Variant) As String
    If Not IsNull(v) Then Texto = Trim(CStr(v))
End Function

Private Function Fecha(ByVal v As Variant) As String
    If IsDate(v) Then Fecha = Format(CDate(v), "dd/mm/yyyy") Else Fecha = Texto(v)
End Function

Private Function Dinero(ByVal v As Variant) As String
    If IsNumeric(v) Then Dinero = Format(CDbl(v), "#,##0.00") Else Dinero = Texto(v)
End Function

Private Function Unir(ByVal a As String, ByVal b As String, ByVal sep As String) As String
    If a <> "" And b <> "" Then Unir = a & sep & b Else Unir = a & b
End Function

' Quita los caracteres que Windows no permite en un nombre de archivo
Private Function NombreArchivo(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "")
    Next ch
    NombreArchivo = Left(Trim(s), 40)
End Function

' Si ya existe una orden con ese nombre, agrega _v2, _v3, ... para no sobrescribirla
Private Function NombreLibre(ByVal base As String) As String
    Dim v As Long
    NombreLibre = base
    v = 1
    Do While Dir(NombreLibre & ".docx") <> "" Or Dir(NombreLibre & ".pdf") <> ""
        v = v + 1
        NombreLibre = base & "_v" & v
    Loop
End Function
