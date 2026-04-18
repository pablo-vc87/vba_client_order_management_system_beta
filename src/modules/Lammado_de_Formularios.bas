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
