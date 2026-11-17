Attribute VB_Name = "Lammado_de_Formularios"


Sub llamar_form_reg()
    ModoFormulario = "REGISTRAR"
    UserForm1.Show
End Sub

Sub llamar_calendario()
    frmCalendario.Show
End Sub

Sub llamar_buscar()
    ModoFormulario = "MODIFICAR"
    Busqueda.Show
End Sub

'Alta de una empresa sin capturar un pedido (asignar a un botón de la hoja Comandos)
Sub llamar_alta_empresa()
    Dim id As Long
    id = frmEmpresa.Registrar("", "", "", "")
    If id > 0 Then
        ThisWorkbook.Save
        MsgBox "Empresa registrada con EmpresaID " & id & ".", vbInformation, "CALEIDO"
    End If
End Sub

'Buscar y eliminar contactos (hoja Clientes)
Sub llamar_contactos()
    frmContactos.Show
End Sub
