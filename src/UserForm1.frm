VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "CALEIDO | Registro de Pedidos"
   ClientHeight    =   4.70472e5
   ClientLeft      =   13920
   ClientTop       =   59496
   ClientWidth     =   1.96380e5
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TotalProductos As Long
Private Eventos As Collection
Private hayCoincidencias As Boolean 'nvo
Private Guardando As Boolean



'*********************************************************************************************************************************
' BOTÓN AGREGAR PRODUCTO
'*********************************************************************************************************************************
Private Sub cmdAgregarProducto_Click()
    
    '=============================
    ' VALIDAR DATOS DEL CLIENTE
    '=============================
    If Not ValidarPagina1 Then Exit Sub
    
    TotalProductos = TotalProductos + 1
    CrearFilaProducto TotalProductos

End Sub

'*********************************************************************************************************************************
' BOTÓN QUITAR PRODUCTO
'*********************************************************************************************************************************

Private Sub cmdQuitarProducto_Click()
    
    '=============================
    ' SEGURIDAD
    '=============================
    If TotalProductos = 0 Then Exit Sub
    
    
    Me.MultiPage1.SetFocus
    
    Dim ctrl As Control
    Dim sufijo As String
    sufijo = CStr(TotalProductos)

    '=============================
    ' ELIMINAR CONTROLES
    '=============================
    For Each ctrl In frProductos.Controls
        If ctrl.Name Like "*" & sufijo Then
            frProductos.Controls.Remove ctrl.Name
        End If
    Next ctrl

    '=============================
    ' AJUSTAR CONTADOR
    '=============================
    TotalProductos = TotalProductos - 1

    '=============================
    ' AJUSTAR SCROLL
    '=============================
    If TotalProductos > 0 Then
        frProductos.ScrollHeight = 210 * TotalProductos + 20
    Else
        frProductos.ScrollHeight = frProductos.InsideHeight
    End If

End Sub

'*********************************************************************************************************************************
'PROCEDIMIENTO PARA CREAR CONTROLES
'*********************************************************************************************************************************
Private Sub CrearFilaProducto(ByVal Index As Long)

    Const ALTO_PRODUCTO As Long = 210
    Dim baseTop As Long
    baseTop = (Index - 1) * ALTO_PRODUCTO + 10

    '========================
    ' TÍTULO PRODUCTO
    '========================
    With frProductos.Controls.Add("Forms.Label.1", "lblProd" & Index)
        .Caption = "Producto " & Index
        .Left = 10
        .Top = baseTop
        .Font.Bold = True
    End With

    '========================
    ' FILA SUPERIOR
    '========================
    CrearCampo "Técnica", "Forms.ComboBox.1", "cmbTec", Index, 10, baseTop + 20, 120
    CrearCampo "Material", "Forms.TextBox.1", "txtMat", Index, 150, baseTop + 20, 120
    CrearCampo "Fecha recepción", "Forms.TextBox.1", "txtFechaRec", Index, 280, baseTop + 20, 110
    CrearCampo "Cantidad", "Forms.TextBox.1", "txtCant", Index, 430, baseTop + 20, 50
    CrearCampo "Precio unitario", "Forms.TextBox.1", "txtPrecio", Index, 490, baseTop + 20, 80

    '========================
    ' CARGAR OPCIONES TÉCNICA
    '========================
    With frProductos.Controls("cmbTec" & Index)
        .AddItem "N/A"
        .AddItem "Serigrafía"
        .AddItem "Serigrafía F y V"
        .AddItem "Bordado"
        .AddItem "Bordado F y V"
        .AddItem "Sublimado"
        .AddItem "Sublimado F y V"
        .AddItem "Impresión Directa"
        .AddItem "Impresión Directa F y V"
        .AddItem "Grabado"
        .AddItem "Grabado F y V"
        .AddItem "Vinil"
        .AddItem "Vinil F y V"
        .AddItem "DTF"
        .AddItem "DTF F y V"
        .Style = fmStyleDropDownList
        .MatchRequired = True
    End With

    '========================
    ' BLOQUEAR ESCRITURA FECHA
    '========================
    With frProductos.Controls("txtFechaRec" & Index)
        .Locked = True
        .BackColor = RGB(240, 240, 240)
        '.Enabled = False
    End With

    '========================
    ' BOTÓN CALENDARIO
    '========================
    With frProductos.Controls.Add("Forms.CommandButton.1", "cmdFechaRec" & Index)
        .Caption = "Cal"
        .Left = 395
        .Top = baseTop + 20
        .Width = 25
        .Height = 20
        .Tag = Index
    End With
    
    '========================
    ' FILA INFERIOR
    '========================
    CrearCampo "Nombre logo", "Forms.TextBox.1", "txtLogo", Index, 10, baseTop + 55, 120
    CrearCampo "Tamaño", "Forms.TextBox.1", "txtTam", Index, 150, baseTop + 55, 120
    CrearCampo "Pantone", "Forms.TextBox.1", "txtPnt", Index, 280, baseTop + 55, 110

    '========================
    ' RUTA IMAGEN (OCULTA)
    '========================
    With frProductos.Controls.Add("Forms.TextBox.1", "txtImg" & Index)
        .Visible = False
    End With

    '========================
    ' IMAGEN PREVIA
    '========================
    With frProductos.Controls.Add("Forms.Image.1", "imgPrev" & Index)
        .Left = 410
        .Top = baseTop + 55
        .Width = 138
        .Height = 144
        .BorderStyle = fmBorderStyleSingle
        .PictureSizeMode = fmPictureSizeModeZoom
    End With

    '========================
    ' BOTÓN IMAGEN
    '========================
    With frProductos.Controls.Add("Forms.CommandButton.1", "cmdImg" & Index)
        .Caption = "Imagen..."
        .Left = 560
        .Top = baseTop + 55
        .Width = 70
        .Tag = Index
    End With

    '========================
    ' AJUSTE SCROLL
    '========================
    frProductos.ScrollHeight = ALTO_PRODUCTO * Index + 20

    '========================
    ' TERCERA FILA – OBSERVACIONES
    '========================
    CrearCampo "Observaciones", "Forms.TextBox.1", "txtObsProd", Index, 10, baseTop + 90, 385

    With frProductos.Controls("txtObsProd" & Index)
        .Multiline = True
        .Height = 40
        .ScrollBars = fmScrollBarsVertical
    End With
    
    '========================
    ' REGISTRAR EVENTOS (CLAVE)
    '========================
    Dim ev As clsEventosProducto
    
    ' Botón calendario
    Set ev = New clsEventosProducto
    Set ev.Btn = frProductos.Controls("cmdFechaRec" & Index)
    Set ev.ParentForm = Me
    ev.Index = Index
    Eventos.Add ev
    
    ' Botón imagen
    Set ev = New clsEventosProducto
    Set ev.Btn = frProductos.Controls("cmdImg" & Index)
    Set ev.ParentForm = Me
    ev.Index = Index
    Eventos.Add ev
    
    ' TextBox cantidad
    Set ev = New clsEventosProducto
    Set ev.Txt = frProductos.Controls("txtCant" & Index)
    Eventos.Add ev
    
    ' TextBox precio
    Set ev = New clsEventosProducto
    Set ev.Txt = frProductos.Controls("txtPrecio" & Index)
    Eventos.Add ev
End Sub

'*********************************************************************************************************************************
'FUNCIÓN AUXILIAR DEL PROCEDIMIENTO ANTERIOR
'*********************************************************************************************************************************
Private Sub CrearCampo(ByVal TextoLabel As String, _
                       ByVal TipoControl As String, _
                       ByVal NombreBase As String, _
                       ByVal Index As Long, _
                       ByVal LeftPos As Long, _
                       ByVal TopPos As Long, _
                       ByVal Ancho As Long)

    Dim lbl As MSForms.Label
    Dim ctrl As Control

    ' LABEL
    Set lbl = frProductos.Controls.Add("Forms.Label.1", "lbl" & NombreBase & Index)
    With lbl
        .Caption = TextoLabel
        .Left = LeftPos
        .Top = TopPos - 12
        .Font.Size = 8
        .ForeColor = RGB(80, 80, 80)
    End With

    ' CONTROL
    Set ctrl = frProductos.Controls.Add(TipoControl, NombreBase & Index)
    With ctrl
        .Left = LeftPos
        .Top = TopPos
        .Width = Ancho
    End With

End Sub

Public Sub CrearProducto(ByVal Index As Long)

    Dim ev As clsEventosProducto

    ' Crear controles visuales
    CrearFilaProducto Index

End Sub

'*********************************************************************************************************************************
'LLENADO DE DATOS DEL FORMULARIO CUANDO SE LLAMA DESDE REGISTRAR O MODIFICAR
'*********************************************************************************************************************************

'CARGA EL NÚMERO DE PEDIDO Y FECHA ACTUAL AUTOMÁTOCAMENTE CUANDO SE VA A REGISTRAR NUEVO PEDIDO
Private Sub UserForm_Activate()

    ' ===== REGISTRO NUEVO =====
    If ModoFormulario = "REGISTRAR" Then
        Me.txtNoPed.Value = NuevoPedidoID()
        Me.txtNoPed.TabStop = False
        Me.txtFechaActual = Date
        ' Forzar foco inicial correcto
        If Me.MultiPage1.Value = 0 Then
            Me.txtNombreContacto.SetFocus
        End If
        Gestionbotones
    End If
    ' ===== MODIFICACIÓN =====
    If ModoFormulario = "MODIFICAR" Then
        If Me.MultiPage1.Value = 0 Then
            Me.txtNombreContacto.SetFocus
        End If
        Gestionbotones
    End If
End Sub

'*********************************************************************************************************************************
'INSTRUCCIONES A LLEVARSE A CABO AL INICIAR EL FORMULARIO
'*********************************************************************************************************************************
Private Sub UserForm_Initialize()
    Guardando = False
    ' EVENTOS COMO COLECCIÓN
    Set Eventos = New Collection
    
    'INICIALIZA CON TOTAL DE PRODUCTOS = 0
    TotalProductos = 0
    
    'PARA QUE FUNCIONEN SCROLLBARS
    With frProductos
        .ScrollBars = fmScrollBarsVertical
        .ScrollTop = 0
    End With
    
    'INICIALIZACIÓN DE LA LISTA DROPDOWN DE CLIENTES
    With Me.lstClientes
        .Visible = False
        .ColumnCount = 1
        .IntegralHeight = False
        .Height = 90
    End With
    
    'MUESTRA EL FORMULARIO CENTRADO EN PANTALLA AL TAMAÑO DESEADO
    UserForm1.StartUpPosition = 2
    UserForm1.Height = 612.6
    UserForm1.Width = 880.2
    frmCalendario.StartUpPosition = 2
    
    'OCULTA LAS PESTAÑAS DE LAS PÁGINAS
    Me.MultiPage1.Style = fmTabStyleNone
    
    
    
    'CONDICIONALES DE CARGA DE LOS COMBOBOX
    If Me.MultiPage1.Pages.Count <> 0 Then
        Me.MultiPage1.Value = 0
    End If
    
    Dim fil, fin, q As Long
    Dim lis As String
    
    fil = Hoja2.Range("C" & Rows.Count).End(xlUp).Row + 1
    fin = fil - 1
    For q = 3 To fin
        lis = Hoja2.Cells(q, 3)
        'CARGA DE DATOS LOS COMBOBOX DE TÉCNICA
        cmbEstatus.AddItem (lis)
    Next q
    'NO PERMITE METER OTRO DATO QUE NO SE ENCUENTRE EN LA LISTA DEL COMBOBOX DE TÉCNICA
    cmbEstatus.Style = fmStyleDropDownList
    'CONDICIÓN QUE EL TELÉFONO NO PUEDE TENER MÁS DE 10 CARACTERES
    txtTel.MaxLength = 10
    txtTel2.MaxLength = 10
    'CAMBIA EL COLOR DE CASILLAS BLOQUEADAS
    txtNoPed.BackColor = &H80000004
    txtFechaActual.BackColor = &H80000004
    
    'EVITA MOSTRAR EL BOTÓN DE GUARDADO AL INICIAR
    cmdGuardar.Visible = False
    
    With Me
    'EVITA INTRODUCCIÓN DEL USUARIO A LOS CAMPOS
    .txtFechaActual.Locked = True
    .txtFechaEntrega.Locked = True
    .txtNoPed.Locked = True
    
    
    End With
    Gestionbotones
End Sub

'*********************************************************************************************************************************
'--------------- FILTRO PARA LISTA CLIENTES ----------------------
'*********************************************************************************************************************************
Private Sub txtNombreContacto_Change()
    
    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim texto As String
    Dim textoLimpio As String

    ' Limpieza de texto (solo letras)
    textoLimpio = LimpiarSoloTexto(Me.txtNombreContacto.Text)

    ' Evitar recursión infinita
    If textoLimpio <> Me.txtNombreContacto.Text Then
        Me.txtNombreContacto.Text = textoLimpio
        Me.txtNombreContacto.SelStart = Len(textoLimpio)
        Exit Sub
    End If

    texto = Trim(textoLimpio)
    Set ws = Sheets("Clientes")
    Me.lstClientes.Clear
    hayCoincidencias = False 'nvo

    ' Menos de 2 letras ? no mostrar dropdown
    If Len(texto) < 2 Then
        Me.lstClientes.Visible = False
        Exit Sub
    End If

    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        If InStr(1, ws.Cells(i, "A").Value, texto, vbTextCompare) > 0 Then
            Me.lstClientes.AddItem ws.Cells(i, "A").Value
        End If
    Next i

    If Me.lstClientes.ListCount > 0 Then
        hayCoincidencias = True 'nvo
        MostrarDropdownClientes
        Me.lstClientes.ListIndex = -1
    Else
        Me.lstClientes.Visible = False
    End If

End Sub

'*********************************************************************************************************************************
'--------------- MUESTRA LISTA COMO DROPBOX ----------------------
'*********************************************************************************************************************************
Private Sub MostrarDropdownClientes()
    With Me.lstClientes
        .Left = Me.txtNombreContacto.Left
        .Top = Me.txtNombreContacto.Top + Me.txtNombreContacto.Height
        .Width = Me.txtNombreContacto.Width
        .Visible = True
        .ZOrder 0
    End With
End Sub

'*********************************************************************************************************************************
'--------------- SELECCIÓN DEL CLIENTE CON RATÓN ----------------------
'*********************************************************************************************************************************
Private Sub lstClientes_Click()

    If Me.lstClientes.ListIndex = -1 Then Exit Sub

    ' Forzar foco primero (clave anti-crash)
    Me.lstClientes.SetFocus

    AutorrellenarCliente Me.lstClientes.Value

End Sub

'*********************************************************************************************************************************
'--------------- BLOQUEO EN TXTNOMBRECONTACTO ----------------------
'*********************************************************************************************************************************
Private Sub txtNombreContacto_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

    Select Case KeyCode
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
            KeyCode = 0   ' Flechas anuladas

        Case vbKeyReturn
            Me.lstClientes.Visible = False
            KeyCode = 0   ' Enter NO selecciona

        Case vbKeyEscape
            Me.lstClientes.Visible = False
            KeyCode = 0   ' Esc cierra

        Case vbKeyTab
            Me.lstClientes.Visible = False
            ' Tab sigue su flujo normal
    End Select

End Sub

'AUTORRELENO DE HOJA 1
Private Sub AutorrellenarCliente(ByVal NombreCliente As String)

    Dim ws As Worksheet
    Dim Fila As Long
    
    Set ws = Sheets("Clientes")
    
    Fila = Application.Match(NombreCliente, ws.Columns("A"), 0)
    
    If IsError(Fila) Then Exit Sub
    
    With ws
        Me.txtNombreContacto.Value = .Cells(Fila, "A").Value
        Me.txtRazonSocial.Value = .Cells(Fila, "B").Value
        Me.txtRFC.Value = .Cells(Fila, "C").Value
        Me.txtTel.Value = .Cells(Fila, "D").Value
        Me.txtTel2.Value = .Cells(Fila, "E").Value
        Me.txtEmail.Value = .Cells(Fila, "F").Value
        Me.txtDomFiscal.Value = .Cells(Fila, "G").Value
    End With
    
    ' Ocultar dropdown
    Me.lstClientes.Visible = False
    
    ' Enfocar siguiente campo lógico
    Me.txtTel.SetFocus
End Sub


'*********************************************************************************************************************************
'--------------- BOTONES DE PÁG SIGUIENTE Y ANTERIOR ----------------------
'*********************************************************************************************************************************
'BOTÓN PÁG ANTERIOR
Private Sub cmdAnterior_Click()
    Me.MultiPage1.Value = Me.MultiPage1.Value - 1
    Me.cmdGuardar.Visible = False
    Gestionbotones
End Sub
'BOTÓN PÁG SIGUIENTE
Private Sub cmdSiguiente_Click()
    ' Quitar foco de controles dinámicos
    Me.MultiPage1.SetFocus
    
    'Solo existe avance de Page 0 -> Page 1
    If Me.MultiPage1.Value = 0 Then
        If Not ValidarPagina1 Then Exit Sub
        Me.MultiPage1.Value = 1
        Gestionbotones
    End If

End Sub

'CAMBIAR PESTAÑA PÁGINA
Private Sub MultiPage1_Change()
    Gestionbotones
End Sub

'CONDICIONALES DEPENDIENDO DEL NÚMERO DE PÁGINAS, ACTIVA Y/O DESACTIVA LOS BOTONES
Private Sub Gestionbotones()

    Dim numpag As Byte
    Dim pagAct As Byte

    numpag = Me.MultiPage1.Pages.Count
    pagAct = Me.MultiPage1.Value

    With Me

        ' Seguridad
        If numpag < 2 Then
            .cmdAnterior.Enabled = False
            .cmdSiguiente.Visible = False
            .cmdGuardar.Visible = True
            Exit Sub
        End If

        ' =========================
        ' PÁGINA 1 – CLIENTE
        ' =========================
        If pagAct = 0 Then
            .cmdAnterior.Enabled = False
            .cmdAnterior.Visible = False
            .cmdSiguiente.Visible = True
            .cmdGuardar.Visible = False

        ' =========================
        ' PÁGINA 2 – PRODUCTOS
        ' =========================
        ElseIf pagAct = 1 Then
            .cmdAnterior.Visible = True
            .cmdAnterior.Enabled = True
            .cmdSiguiente.Visible = False
            .cmdGuardar.Visible = True
        End If

        .lblpagXdeY.Caption = "Página " & pagAct + 1 & " de " & numpag

    End With
End Sub

'*********************************************************************************************************************************
'------------------------ BOTONOES SOLO NÚMERO -------------------------------------------
'*********************************************************************************************************************************
'VALIDA QUE EL TELEFONO SEAN SOLO NÚMEROS
Private Sub txtTel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub
Private Sub txtTel2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

'*********************************************************************************************************************************
'------------ MSJ DE CORREO MAL ESCRITO -----------------------------
'*********************************************************************************************************************************

Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Valida_Correo(txtEmail) = False Then
        MsgBox "Escriba una dirección de correo electrónico válida", , "CALEIDO"
        Cancel = True
    Else
    End If
End Sub

'*********************************************************************************************************************************
'--------------- FUNCIONES PARA EL BOTÓN GUARDAR ---------------
'*********************************************************************************************************************************
' ====================================
' FUNCIÓN BUSCAR CLIENTES POR NOMBRES------NUEVO
' ====================================
Private Function FilaCliente(ByVal NombreCliente As String) As Variant
    Dim ws As Worksheet
    Dim Fila As Variant

    Set ws = Sheets("Clientes")

    Fila = Application.Match(NombreCliente, ws.Columns("A"), 0)

    If IsError(Fila) Then
        FilaCliente = 0
    Else
        FilaCliente = Fila
    End If
End Function

' ====================================
' FUNCIÓN COMPARAR INFORMACIÓN DEL CLIENTE------NUEVO
' ====================================
Private Function ClienteEsIgual(ByVal Fila As Long) As Boolean
    With Sheets("Clientes")
        ClienteEsIgual = _
            Trim(.Cells(Fila, "B").Value) = Trim(Me.txtRazonSocial.Value) And _
            Trim(.Cells(Fila, "C").Value) = Trim(Me.txtRFC.Value) And _
            Trim(.Cells(Fila, "D").Value) = Trim(Me.txtTel.Value) And _
            Trim(.Cells(Fila, "E").Value) = Trim(Me.txtTel2.Value) And _
            Trim(.Cells(Fila, "F").Value) = Trim(Me.txtEmail.Value) And _
            Trim(.Cells(Fila, "G").Value) = Trim(Me.txtDomFiscal.Value)
    End With
End Function



' ====================================
' PROCEDIMIENTO GUARDAR PEDIDO y Detalles------NUEVO
' ====================================
Private Sub GuardarPedido()
    Dim wsPed As Worksheet
    Dim wsDet As Worksheet
    Dim filaPed As Long
    Dim filaDet As Long
    Dim PedidoID As Long
    Dim subtotalPed As Double
    Dim ivaPed As Double
    Dim totalPed As Double
    Dim subtotalDet As Double
    Dim ivaDet As Double
    Dim totalDet As Double
    Dim i As Long
    

    Set wsPed = Sheets("Pedidos")
    Set wsDet = Sheets("Detalle_Pedidos")
    filaPed = wsPed.Cells(wsPed.Rows.Count, 1).End(xlUp).Row + 1
    PedidoID = NuevoPedidoID()
    
    '====================
    ' GUARDAR PEDIDO
    '====================

    With wsPed
        .Cells(filaPed, "A").Value = PedidoID
        .Cells(filaPed, "B").Value = Date
        .Cells(filaPed, "C").Value = Me.txtNombreContacto.Value 'nombrecliente
        .Cells(filaPed, "D").Value = Me.txtRazonSocial.Value 'bien
        .Cells(filaPed, "E").Value = Me.txtRFC.Value 'RFC
        .Cells(filaPed, "F").Value = Me.txtTel.Value 'Telefono
        .Cells(filaPed, "G").Value = Me.txtTel2.Value 'telefono2
        .Cells(filaPed, "H").Value = Me.txtEmail.Value 'email
        .Cells(filaPed, "I").Value = Me.txtDomFiscal.Value 'domfiscal
        .Cells(filaPed, "J").Value = Me.txtDirEntrega.Value 'Direccion entrega
        .Cells(filaPed, "K").Value = Me.txtFechaEntrega.Value 'bien
        .Cells(filaPed, "L").Value = Me.cmbEstatus.Value 'Estatus
        
    End With
    
    '====================
    ' GUARDAR PRODUCTOS
    '====================
    For i = 1 To TotalProductos

        filaDet = wsDet.Cells(wsDet.Rows.Count, 1).End(xlUp).Row + 1

        subtotalDet = Val(frProductos.Controls("txtCant" & i).Value) * _
           Val(frProductos.Controls("txtPrecio" & i).Value)
        ivaDet = subtotalDet * 0.16
        totalDet = subtotalDet + ivaDet
        With wsDet
            .Cells(filaDet, 1).Value = PedidoID
            .Cells(filaDet, 2).Value = i
            .Cells(filaDet, 3).Value = frProductos.Controls("cmbTec" & i).Value
            .Cells(filaDet, 4).Value = frProductos.Controls("txtMat" & i).Value
            .Cells(filaDet, 5).Value = frProductos.Controls("txtFechaRec" & i).Value
            .Cells(filaDet, 6).Value = frProductos.Controls("txtCant" & i).Value
            .Cells(filaDet, 7).Value = frProductos.Controls("txtPrecio" & i).Value
            .Cells(filaDet, 8).Value = frProductos.Controls("txtLogo" & i).Value
            .Cells(filaDet, 9).Value = frProductos.Controls("txtTam" & i).Value
            .Cells(filaDet, 10).Value = frProductos.Controls("txtPnt" & i).Value
            .Cells(filaDet, 11).Value = frProductos.Controls("txtImg" & i).Value
            .Cells(filaDet, 12).Value = frProductos.Controls("txtObsProd" & i).Value
            .Cells(filaDet, 13).Value = subtotalDet
            .Cells(filaDet, 14).Value = ivaDet
            .Cells(filaDet, 15).Value = totalDet
        End With

        totalPed = totalPed + totalDet
        subtotalPed = subtotalPed + subtotalDet
        ivaPed = ivaPed + ivaDet

    Next i
    
    With wsPed
        .Cells(filaPed, "M").Value = subtotalPed
        .Cells(filaPed, "N").Value = ivaPed
        .Cells(filaPed, "O").Value = totalPed
    End With
End Sub



'*********************************************************************************************************************************
'--------------- BOTÓN GUARDAR -------------------------------------------------------------
'*********************************************************************************************************************************
'GUARDA LOS DATOS INGRESADOS EN LA BASE DE DATOS
Private Sub cmdGuardar_Click()
    
    '=========================
    ' BLOQUEO ANTI DOBLE CLICK
    '=========================
    If Guardando Then Exit Sub
    Guardando = True
    Me.cmdGuardar.Enabled = False

    On Error GoTo SalidaSegura

    '=========================
    ' VALIDACIONES
    '=========================
    If Not ValidarPagina1 Then GoTo SalidaSegura
    If Not ValidarPagina2 Then GoTo SalidaSegura
    Application.ScreenUpdating = False

    '=========================
    ' REGISTRAR / MODIFICAR
    '=========================
    If ModoFormulario = "REGISTRAR" Then
        
        GuardarPedido
        MsgBox "Pedido guardado correctamente", vbInformation, "CALEIDO"
        
    ElseIf ModoFormulario = "MODIFICAR" Then

        ActualizarPedido
        EliminarDetallePedido PedidoID_Activo
        GuardarDetallePedido PedidoID_Activo
        RecalcularTotalesPedido PedidoID_Activo

        MsgBox "Pedido actualizado correctamente", vbInformation, "CALEIDO"

    End If
    '=========================
    ' DECISIÓN ÚNICA SOBRE CLIENTES (ETAPA 2)
    '=========================
    Dim rCliente As ResultadoCliente
    rCliente = DetectarClienteExistente( _
    Me.txtNombreContacto.Value, _
    Me.txtRFC.Value, _
    Me.txtTel.Value, _
    Me.txtTel2.Value, _
    Me.txtEmail.Value _
    )

    DecidirActualizacionCliente rCliente
    
    ThisWorkbook.Save
    Call ReiniciarFormulario

SalidaSegura:
    Guardando = False
    Me.cmdGuardar.Enabled = True
    Application.ScreenUpdating = True
End Sub


'*********************************************************************************************************************************
'--------- BOTONES DE COMANDO DE FECHAS ------------------------------------
'*********************************************************************************************************************************

'FECHA DE ENTREGA
Private Sub cmdFechaEntrega_Click()
    Set ModuloCalendario.TextBoxDestino = Me.txtFechaEntrega
    frmCalendario.Show
End Sub

'*********************************************************************************************************************************
'--------- REINICIO FORMULARIO DESPUES DE GUARDAR ------------------------------------
'*********************************************************************************************************************************
Private Sub ReiniciarFormulario()

    Dim ctrl As Control

    '=========================
    ' REINICIAR PÁGINA 1
    '=========================
    Me.txtNombreContacto.Value = ""
    Me.txtRazonSocial.Value = ""
    Me.txtRFC.Value = ""
    Me.txtTel.Value = ""
    Me.txtTel2.Value = ""
    Me.txtEmail.Value = ""
    Me.txtDomFiscal.Value = ""
    Me.txtDirEntrega.Value = ""
    Me.cmbEstatus.Value = ""

    '=========================
    ' REGENERAR DATOS AUTOMÁTICOS
    '=========================
    Me.txtNoPed.Value = NuevoPedidoID()
    Me.txtFechaActual.Value = Date

    '=========================
    ' LIMPIAR PRODUCTOS DINÁMICOS
    '=========================
    For Each ctrl In Me.frProductos.Controls
        Me.frProductos.Controls.Remove ctrl.Name
    Next ctrl

    '=========================
    ' REINICIAR CONTADORES
    '=========================
    TotalProductos = 0
    Me.frProductos.ScrollHeight = Me.frProductos.InsideHeight
    LimpiarImagenesProductos
    
    ' ==========================
    ' LIMPIAR EVENTOS
    ' ==========================
    Set Eventos = Nothing
    Set Eventos = New Collection
    
    
    '=========================
    ' VOLVER A PÁGINA 1
    '=========================
    Me.MultiPage1.Value = 0
    Me.txtNombreContacto.SetFocus
    ModoFormulario = "REGISTRAR"

End Sub

'===================

'===================

'===================
Public Sub RenderizarImagenesProductos()

    Dim ctrl As Control
    Dim idx As Long
    Dim ruta As String
    Dim img As MSForms.Image
    
    If Me Is Nothing Then Exit Sub
    If Not Me.Visible Then Exit Sub
    
    For Each ctrl In Me.frProductos.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            If ctrl.Name Like "txtImg*" Then

                ruta = ctrl.Value
                idx = Val(Replace(ctrl.Name, "txtImg", ""))
                
                On Error Resume Next
                Set img = Me.frProductos.Controls("imgPrev" & idx)
                On Error GoTo 0

                If Not img Is Nothing Then
                    If ruta <> "" And Dir(ruta) <> "" Then
                        img.Picture = LoadPicture(ruta)
                    Else
                        img.Picture = Nothing
                    End If
                End If

            End If
        End If
    Next ctrl

End Sub
'===================

'*********************************************************************************************************************************
'--------- PROCEDIMIENTO CARGA PEDIDO EN MODIFICAR ------------------------------------
'*********************************************************************************************************************************
'Private Sub CargarPedido(ByVal PedidoID As Long)
    
    'CargandoPedido = True
    
'    Dim wsPed As Worksheet
'    Dim wsDet As Worksheet
'    Dim fila As Range

'    Set wsPed = Sheets("Pedidos")
'    Set wsDet = Sheets("Detalle_Pedidos")

    ' Buscar pedido
'    Set fila = wsPed.Columns("A").Find(PedidoID, LookAt:=xlWhole)

'    If fila Is Nothing Then
'        MsgBox "Pedido no encontrado", vbCritical, "CALEIDO"
'        Exit Sub
'    End If

    ' ===== DATOS DEL PEDIDO / CLIENTE =====
'    With Me
'        .txtNoPed.Value = PedidoID
'        .txtFechaActual.Value = fila.Offset(0, 1).Value
'        .txtNombreContacto.Value = fila.Offset(0, 2).Value
'        .txtRazonSocial.Value = fila.Offset(0, 3).Value
'        .txtRFC.Value = fila.Offset(0, 4).Value
'        .txtTel.Value = fila.Offset(0, 5).Value
'        .txtTel2.Value = fila.Offset(0, 6).Value
'        .txtEmail.Value = fila.Offset(0, 7).Value
'        .txtDomFiscal.Value = fila.Offset(0, 8).Value
'        .txtDirEntrega.Value = fila.Offset(0, 9).Value
'        .txtFechaEntrega.Value = fila.Offset(0, 10).Value
'        .cmbEstatus.Value = fila.Offset(0, 11).Value
'    End With

    ' ===== DETALLE =====
'    Call CargarDetallePedido(PedidoID)

'End Sub

'*********************************************************************************************************************************
'--------- PROCEDIMIENTO LIMPIAR PRODUCTOS ------------------------------------
'*********************************************************************************************************************************
Public Sub LimpiarProductos()

    Dim ctrl As Control

    For Each ctrl In Me.frProductos.Controls
        Me.frProductos.Controls.Remove ctrl.Name
    Next ctrl
    ' ==========================
    ' LIMPIAR EVENTOS
    ' ==========================
    TotalProductos = 0
    frProductos.ScrollHeight = frProductos.InsideHeight

End Sub

Public Sub LimpiarImagenesProductos()

    Dim ctrl As Control

    For Each ctrl In Me.frProductos.Controls
        If TypeName(ctrl) = "Image" Then
            ctrl.Picture = Nothing
        End If
    Next ctrl

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    LimpiarImagenesProductos
End Sub


' ============================================
' ---------- FLECHA ABAJO ENTRA A LISTA -----------
' ============================================
Private Sub lstClientes_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

    KeyCode = 0   ' Teclado totalmente deshabilitado

End Sub

' ============================================
' ---------- CIERRA LISTA AL CLICK EN OTRO LADO -----------
' ============================================
Private Sub OcultarListaClientes()
    Me.lstClientes.Visible = False
End Sub

' ============================================
' ---------- CONTROLES EN TXTBX PARA OCULTAR LISTA -----------
' ============================================
Private Sub txtRazonSocial_Enter()
    OcultarListaClientes
End Sub

Private Sub txtRFC_Enter()
    OcultarListaClientes
End Sub

Private Sub txtTel_Enter()
    OcultarListaClientes
End Sub

Private Sub txtEmail_Enter()
    OcultarListaClientes
End Sub

Private Sub txtDomFiscal_Enter()
    OcultarListaClientes
End Sub

Private Sub txtDirEntrega_Enter()
    OcultarListaClientes
End Sub

Private Sub cmbEstatus_Enter()
    OcultarListaClientes
End Sub




' ====================================
' PROCEDIMIENTO GUARDAR CLIENTE------NUEVO
' ====================================
Public Sub GuardarCliente()

    Dim ws As Worksheet
    Dim Fila As Long
    Dim opcion As VbMsgBoxResult

    Set ws = Sheets("Clientes")

    Fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    With ws
        .Cells(Fila, "A").Value = Trim(Me.txtNombreContacto.Value)
        .Cells(Fila, "B").Value = Trim(Me.txtRazonSocial.Value)
        .Cells(Fila, "C").Value = Trim(Me.txtRFC.Value)
        .Cells(Fila, "D").Value = Trim(Me.txtTel.Value)
        .Cells(Fila, "E").Value = Trim(Me.txtTel2.Value)
        .Cells(Fila, "F").Value = Trim(Me.txtEmail.Value)
        .Cells(Fila, "G").Value = Trim(Me.txtDomFiscal.Value)
    End With

End Sub

'=================================================
' PROCESO ACTUALIZAR CLIENTE
'=================================================
Public Sub ActualizarCliente(ByVal FilaCliente As Long)

    Dim ws As Worksheet
    Dim celda As Range
    Dim Fila As Long

    Set ws = Sheets("Clientes")

    ' 1?? Buscar primero por RFC (si existe)
    If Trim(UserForm1.txtRFC.Value) <> "" Then
        Set celda = ws.Columns("C").Find( _
            Trim(UserForm1.txtRFC.Value), LookAt:=xlWhole)
    End If

    ' 2?? Si no se encontró por RFC, buscar por nombre
    If celda Is Nothing Then
        Set celda = ws.Columns("A").Find( _
            Trim(UserForm1.txtNombreContacto.Value), LookAt:=xlWhole)
    End If

    ' 3?? Si no existe, salir
    If celda Is Nothing Then Exit Sub

    ' 4?? OBTENER NÚMERO DE FILA REAL
    Fila = celda.Row

    ' 5?? ESCRIBIR SIEMPRE DESDE LA HOJA
    With Sheets("Clientes")
    .Cells(FilaCliente, "A").Value = Trim(txtNombreContacto.Value)
    .Cells(FilaCliente, "B").Value = Trim(txtRazonSocial.Value)
    .Cells(FilaCliente, "C").Value = Trim(txtRFC.Value)
    .Cells(FilaCliente, "D").Value = Trim(txtTel.Value)
    .Cells(FilaCliente, "E").Value = Trim(txtTel2.Value)
    .Cells(FilaCliente, "F").Value = Trim(txtEmail.Value)
    .Cells(FilaCliente, "G").Value = Trim(txtDomFiscal.Value)
End With

End Sub

