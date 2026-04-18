VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Eliminar 
   Caption         =   "CALEIDO | Eliminar"
   ClientHeight    =   1956
   ClientLeft      =   -180
   ClientTop       =   -696
   ClientWidth     =   2544
   OleObjectBlob   =   "Eliminar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Eliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelEliminar_Click()
    Unload Me
End Sub

'*********************************************************************************************************************************
'--------------- BOTÓN ACEPTAR ELIMINAR ----------------------
'*********************************************************************************************************************************
Private Sub cmdEliminarReg_Click()

    If txtPswdEliminar.Value <> "12345" Then
        MsgBox "Contraseña incorrecta", vbCritical, "CALEIDO"
        txtPswdEliminar.Value = ""
        txtPswdEliminar.SetFocus
        Exit Sub
    End If

    If PedidoID_Activo = 0 Then
        MsgBox "No hay pedido seleccionado", vbCritical, "CALEIDO"
        Exit Sub
    End If

    Call EliminarPedido(PedidoID_Activo)

    MsgBox "Se eliminó el pedido: " & PedidoID_Activo, vbInformation, "CALEIDO"

    Unload Me
    Busqueda.cmdBuscar_Click

End Sub


Private Sub UserForm_Activate()
    Me.Height = 220.8
    Me.Width = 322.8
    Me.StartUpPosition = 2
    
    lblEliminar.Caption = "¿Seguro que quiere eliminar pedido " & PedidoID_Activo & "?"
    txtPswdEliminar.Value = ""
    txtPswdEliminar.PasswordChar = "*"
End Sub

'Private Sub cmdAceptar_Click()

'    If Me.txtPassword.Value <> "tu_password" Then
'        MsgBox "Contraseña incorrecta", vbCritical
'        Exit Sub
'    End If

'    Call EliminarPedido PedidoID_Activo

'    MsgBox "Pedido eliminado correctamente", vbInformation

'    Unload Me
'End Sub

