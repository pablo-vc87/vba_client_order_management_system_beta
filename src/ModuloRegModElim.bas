Attribute VB_Name = "ModuloRegModElim"
Public ModoFormulario As String
Public PedidoID_Activo As Long



'=================================================
' PROCESO ELIMINAR PEDIDO
'=================================================
Public Sub EliminarPedido(ByVal PedidoID As Long)

    Dim wsPed As Worksheet
    Dim wsDet As Worksheet
    Dim i As Long

    Set wsPed = Sheets("Pedidos")
    Set wsDet = Sheets("Detalle_Pedidos")

    'Eliminar detalle
    For i = wsDet.Cells(wsDet.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If wsDet.Cells(i, "A").Value = PedidoID Then
            wsDet.Rows(i).Delete
        End If
    Next i

    'Eliminar pedido
    For i = wsPed.Cells(wsPed.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If wsPed.Cells(i, "A").Value = PedidoID Then
            wsPed.Rows(i).Delete
            Exit For
        End If
    Next i

    'ThisWorkbook.Save

End Sub



'=================================================
' PROCESO ACTUALIZAR PEDIDO
'=================================================
Public Sub ActualizarPedido()

    Dim ws As Worksheet
    Dim Fila As Range

    Set ws = Sheets("Pedidos")

    Set Fila = ws.Columns("A").Find(PedidoID_Activo, LookAt:=xlWhole)

    If Fila Is Nothing Then
        MsgBox "Pedido no encontrado", vbCritical
        Exit Sub
    End If

    With Fila
        .Cells(1, "B").Value = UserForm1.txtFechaActual.Value
        .Cells(1, "C").Value = UserForm1.txtNombreContacto.Value
        .Cells(1, "D").Value = UserForm1.txtRazonSocial.Value
        .Cells(1, "E").Value = UserForm1.txtRFC.Value
        .Cells(1, "F").Value = UserForm1.txtTel.Value
        .Cells(1, "G").Value = UserForm1.txtTel2.Value
        .Cells(1, "H").Value = UserForm1.txtEmail.Value
        .Cells(1, "I").Value = UserForm1.txtDomFiscal.Value
        .Cells(1, "J").Value = UserForm1.txtDirEntrega.Value
        .Cells(1, "K").Value = UserForm1.txtFechaEntrega.Value
        .Cells(1, "L").Value = UserForm1.cmbEstatus.Value
        ' M, N, O se calculan internamente
    End With

End Sub

'=================================================
' PROCESO ACTUALIZAR DETALLE PEDIDO
'=================================================
Public Sub ActualizarDetallePedido()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = Sheets("Detalle_Pedidos")

    'Eliminar detalle anterior
    For i = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If ws.Cells(i, "A").Value = PedidoID_Activo Then
            ws.Rows(i).Delete
        End If
    Next i

    'Insertar nuevo detalle
    Call GuardarDetalleNuevo(PedidoID_Activo)

End Sub

'=================================================
' ENCUENTRA EL PEDIDO A CARGAR
'=================================================
Public Function FilaPedido(ByVal PedidoID As Long) As Long

    Dim ws As Worksheet
    Dim f As Range

    Set ws = Sheets("Pedidos")
    Set f = ws.Columns("A").Find(PedidoID, LookAt:=xlWhole)

    If f Is Nothing Then
        FilaPedido = 0
    Else
        FilaPedido = f.Row
    End If

End Function

'=================================================
' PROCEDIMIENTO CARGAR PEDIDO (PRINCIPAL ORQUESTADOR)
'=================================================
Public Sub CargarPedido(ByVal PedidoID As Long)
    
    
    Dim Fila As Long
    Fila = FilaPedido(PedidoID)
    
    If Fila = 0 Then
        MsgBox "El pedido no existe", vbCritical, "CALEIDO"
        Exit Sub
    End If
    
    '===================
    ' Avisamos al formulario que viene una carga grande
    
    With UserForm1
        .LimpiarProductos

        CargarPedidoDatos Fila
        CargarClienteDatos Fila
        CargarDetallePedido PedidoID

        .Show   ' ?? AQUÍ VA, y SOLO AQUÍ
    End With
    '===================
End Sub

'=================================================
' CARGAR DATOS DE PEDIDO (vacía datos de pedido en formulario)
'=================================================
Public Sub CargarPedidoDatos(ByVal Fila As Long)

    Dim ws As Worksheet
    Set ws = Sheets("Pedidos")

    With UserForm1
        .txtNoPed.Value = ws.Cells(Fila, "A").Value
        .txtFechaActual.Value = ws.Cells(Fila, "B").Value
        .txtDirEntrega.Value = ws.Cells(Fila, "J").Value
        .txtFechaEntrega.Value = ws.Cells(Fila, "K").Value
        .cmbEstatus.Value = ws.Cells(Fila, "L").Value
    End With

End Sub

'=================================================
' CARGAR (vacía) DATOS DE CLIENTE DESDE HOJA DE PEDIDOS
'=================================================
Public Sub CargarClienteDatos(ByVal Fila As Long)

    Dim ws As Worksheet
    Set ws = Sheets("Pedidos")

    With UserForm1
        .txtNombreContacto.Value = ws.Cells(Fila, "C").Value
        .txtRazonSocial.Value = ws.Cells(Fila, "D").Value
        .txtRFC.Value = ws.Cells(Fila, "E").Value
        .txtTel.Value = ws.Cells(Fila, "F").Value
        .txtTel2.Value = ws.Cells(Fila, "G").Value
        .txtEmail.Value = ws.Cells(Fila, "H").Value
        .txtDomFiscal.Value = ws.Cells(Fila, "I").Value
    End With

End Sub

'=================================================
' ORQUESTADOR DETALLE PEDIDOS
'=================================================
Public Sub CargarDetallePedido(ByVal PedidoID As Long)
    
    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long

    Set ws = Sheets("Detalle_Pedidos")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        If ws.Cells(i, "A").Value = PedidoID Then
            CargarUnProducto ws, i
        End If
    Next i
End Sub

'=================================================
' CREADOR DINÁMICO DE PRODUCTOS
'=================================================
Public Sub CargarUnProducto(ws As Worksheet, ByVal Fila As Long)

    Dim idx As Long
    Dim rutaImg As String
    Dim img As MSForms.Image
    
    idx = UserForm1.TotalProductos + 1
    UserForm1.TotalProductos = idx

    'CREAR CONTROLES
    UserForm1.CrearProducto idx
    

    With UserForm1.frProductos
        .Controls("cmbTec" & idx).Value = ws.Cells(Fila, "C").Value
        .Controls("txtMat" & idx).Value = ws.Cells(Fila, "D").Value
        .Controls("txtFechaRec" & idx).Value = ws.Cells(Fila, "E").Value
        .Controls("txtCant" & idx).Value = ws.Cells(Fila, "F").Value
        .Controls("txtPrecio" & idx).Value = ws.Cells(Fila, "G").Value
        .Controls("txtLogo" & idx).Value = ws.Cells(Fila, "H").Value
        .Controls("txtTam" & idx).Value = ws.Cells(Fila, "I").Value
        .Controls("txtPnt" & idx).Value = ws.Cells(Fila, "J").Value
        .Controls("txtImg" & idx).Value = ws.Cells(Fila, "K").Value
        .Controls("txtObsProd" & idx).Value = ws.Cells(Fila, "L").Value
        ' ========================
        ' CARGA SEGURA DE IMAGEN
        ' ========================

        rutaImg = ws.Cells(Fila, "K").Value
        
        Set img = .Controls("imgPrev" & idx)
        
        
        'img.Picture = LoadPicture(rutaImg)
        Call CargarImagenSeguro(img, rutaImg)
    End With
End Sub

'=================================================
' AL MODIFICAR, BORRA DETALLES PEDIDO ANTES DE GUARDAR
'=================================================
Public Sub EliminarDetallePedido(ByVal PedidoID As Long)

    Dim ws As Worksheet
    Dim i As Long

    Set ws = Sheets("Detalle_Pedidos")

    For i = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If ws.Cells(i, "A").Value = PedidoID Then
            ws.Rows(i).Delete
        End If
    Next i

End Sub

'=================================================
' GUARDA DETALLES EN ORDEN CORRETO DEPUÉS DE BORRAR (PASO ANTERIOR)
'=================================================
Public Sub GuardarDetallePedido(ByVal PedidoID As Long)

    Dim ws As Worksheet
    Dim Fila As Long
    Dim i As Long
    Dim subtotalDet As Double, ivaDet As Double, totalDet As Double
    
    Set ws = Sheets("Detalle_Pedidos")
    

    For i = 1 To UserForm1.TotalProductos
        Fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        subtotalDet = Val(UserForm1.frProductos.Controls("txtCant" & i).Value) * _
                      Val(UserForm1.frProductos.Controls("txtPrecio" & i).Value)
        ivaDet = subtotalDet * 0.16
        totalDet = subtotalDet + ivaDet
        With ws
            .Cells(Fila, "A").Value = PedidoID
            .Cells(Fila, "B").Value = i
            .Cells(Fila, "C").Value = UserForm1.frProductos.Controls("cmbTec" & i).Value
            .Cells(Fila, "D").Value = UserForm1.frProductos.Controls("txtMat" & i).Value
            .Cells(Fila, "E").Value = UserForm1.frProductos.Controls("txtFechaRec" & i).Value
            .Cells(Fila, "F").Value = Val(UserForm1.frProductos.Controls("txtCant" & i).Value)
            .Cells(Fila, "G").Value = Val(UserForm1.frProductos.Controls("txtPrecio" & i).Value)
            .Cells(Fila, "H").Value = UserForm1.frProductos.Controls("txtLogo" & i).Value
            .Cells(Fila, "I").Value = UserForm1.frProductos.Controls("txtTam" & i).Value
            .Cells(Fila, "J").Value = UserForm1.frProductos.Controls("txtPnt" & i).Value
            .Cells(Fila, "K").Value = UserForm1.frProductos.Controls("txtImg" & i).Value
            .Cells(Fila, "L").Value = UserForm1.frProductos.Controls("txtObsProd" & i).Value
            .Cells(Fila, "M").Value = subtotalDet
            .Cells(Fila, "N").Value = ivaDet
            .Cells(Fila, "O").Value = totalDet
        End With
    Next i

End Sub

'=================================================
' Recalcula los totales
'=================================================
Public Sub RecalcularTotalesPedido(ByVal PedidoID As Long)

    Dim wsDet As Worksheet, wsPed As Worksheet
    Dim i As Long
    Dim subT As Double, iva As Double

    Set wsDet = Sheets("Detalle_Pedidos")
    Set wsPed = Sheets("Pedidos")

    For i = 2 To wsDet.Cells(wsDet.Rows.Count, "A").End(xlUp).Row
        If wsDet.Cells(i, "A").Value = PedidoID Then
            subT = subT + wsDet.Cells(i, "M").Value
            iva = iva + wsDet.Cells(i, "N").Value
        End If
    Next i

    Dim filaPed As Range
    Set filaPed = wsPed.Columns("A").Find(PedidoID, LookAt:=xlWhole)

    If Not filaPed Is Nothing Then
        filaPed.Offset(0, 12).Value = subT
        filaPed.Offset(0, 13).Value = iva
        filaPed.Offset(0, 14).Value = subT + iva
    End If

End Sub

Public Sub CargarImagenSeguro(ByVal img As MSForms.Image, ByVal ruta As String)
    
    On Error Resume Next
    
    img.Picture = LoadPicture(ruta)
    
    If Err.Number <> 0 Then
        Err.Clear
        img.Picture = Nothing
    End If
    
    On Error GoTo 0
End Sub
