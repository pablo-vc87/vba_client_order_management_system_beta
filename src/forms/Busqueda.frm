VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda 
   Caption         =   "CALEIDO | Búsqueda"
   ClientHeight    =   71388.01
   ClientLeft      =   5364
   ClientTop       =   22932
   ClientWidth     =   1.96380e5
   OleObjectBlob   =   "Busqueda.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*********************************************************************************************************************************
'BOTÓN QUE JALA LA INFO DE UN PEDIDO EXISTENTE CUANDO SE VA A MODIFICAR
'*********************************************************************************************************************************
Private Sub cmdModificar_Click()
    If Me.lstbxBuscar.ListIndex = -1 Then
        MsgBox "No ha seleccionado ningún registro", vbExclamation, "CALEIDO"
        Exit Sub
    End If
    
    ModoFormulario = "MODIFICAR"
    PedidoID_Activo = CLng(Me.lstbxBuscar.List(Me.lstbxBuscar.ListIndex, 0))
    
    Unload Me
    Call CargarPedido(PedidoID_Activo)
End Sub


'CARACTERÍSTICAS DEL FORMULARIO AL INICIAR
Private Sub UserForm_Initialize()

    Busqueda.Height = 480.6
    Busqueda.Width = 919.8

    Dim ws As Worksheet
    Dim Fila As Long

    Set ws = Sheets("Pedidos")
    cmdBuscar.Visible = False
    With Me.lstbxBuscar
        .Clear
        .ColumnCount = 9
    End With

    For Fila = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Me.lstbxBuscar.AddItem ws.Cells(Fila, "A").Value ' PedidoID
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 1) = ws.Cells(Fila, "B").Value ' Fecha
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 2) = ws.Cells(Fila, "C").Value ' Cliente
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 3) = ws.Cells(Fila, "F").Value ' Telefono
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 4) = ws.Cells(Fila, "H").Value ' Email
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 5) = ws.Cells(Fila, "D").Value ' Razón Social
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 6) = ws.Cells(Fila, "L").Value ' Estatus
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 7) = ws.Cells(Fila, "J").Value ' Direccion
        Me.lstbxBuscar.List(Me.lstbxBuscar.ListCount - 1, 8) = ws.Cells(Fila, "O").Value ' Total
    Next Fila
End Sub

Private Sub txtBuscar_Change()
    cmdBuscar_Click
End Sub

'BOTÓN DE COMANDO PARA BUSCAR
Public Sub cmdBuscar_Click()

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim Criterio As String

    Application.ScreenUpdating = False
    Set ws = Sheets("Pedidos")

    lstbxBuscar.Clear
    Criterio = UCase(Trim(txtBuscar.Value))

    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila

        If Criterio = "" Or _
           UCase(ws.Cells(i, "C").Value) Like "*" & Criterio & "*" Or _
           UCase(ws.Cells(i, "E").Value) Like "*" & Criterio & "*" Or _
           UCase(ws.Cells(i, "D").Value) Like "*" & Criterio & "*" Then

            lstbxBuscar.AddItem
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 0) = ws.Cells(i, "A").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 1) = ws.Cells(i, "B").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 2) = ws.Cells(i, "C").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 3) = ws.Cells(i, "F").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 4) = ws.Cells(i, "H").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 5) = ws.Cells(i, "D").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 6) = ws.Cells(i, "L").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 7) = ws.Cells(i, "J").Value
            lstbxBuscar.List(lstbxBuscar.ListCount - 1, 8) = ws.Cells(i, "O").Value

        End If
    Next i
    Application.ScreenUpdating = True
End Sub



'DOBLE CLICK EN OBJETO DE LA LISTA BUSCAR
'Private Sub lstbxBuscar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'    Dim Cod_Ped As String
'    Cod_Ped = Me.lstbxBuscar.List(Me.lstbxBuscar.ListIndex, 0)
'    MsgBox Cod_Ped, , "CALEIDO"

'End Sub
'AL DAR CLICK EN ELIMINAR
Private Sub cmdEliminar_Click()
'ASEGURANDO QUE EL USUARIO HAYA SELECCIONADO ALGO DE LA LIST BOX
    If Busqueda.lstbxBuscar.ListIndex = -1 Then
        MsgBox "No ha seleccionado ningún registro", vbExclamation, "CALEIDO"
        Exit Sub
    End If
    
    ' Definir modo y pedido activo
    ModoFormulario = "ELIMINAR"
    PedidoID_Activo = CLng(Me.lstbxBuscar.List(Me.lstbxBuscar.ListIndex, 0))

    ' Mostrar formulario de contraseña
    Eliminar.Show
End Sub
