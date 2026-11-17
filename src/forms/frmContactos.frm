VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContactos 
   Caption         =   "CALEIDO | Contactos"
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmContactos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Contactos (hoja Clientes): buscar, modificar y eliminar.
' Los pedidos guardan su propia copia de los datos del cliente, así que modificar o
' eliminar un contacto no cambia pedidos ni órdenes ya generadas.
' Los controles se crean al abrir el formulario: no hay que dibujarlos.

Private mFila As Long                'fila del contacto seleccionado en Clientes (0 = ninguno)

Private WithEvents txtBuscarContacto As MSForms.TextBox
Private WithEvents lstContactos As MSForms.ListBox
Private WithEvents cmdGuardarContacto As MSForms.CommandButton
Private WithEvents cmdEditarEmpresa As MSForms.CommandButton
Private WithEvents cmdEliminarContacto As MSForms.CommandButton
Private WithEvents cmdCerrarContactos As MSForms.CommandButton
Private cmbEmpresaContacto As MSForms.ComboBox

Private Sub UserForm_Initialize()
    Me.Caption = "CALEIDO | Contactos"
    Me.Width = 720
    Me.Height = 430

    Etiqueta "Buscar (nombre, empresa, teléfono o correo):", 12, 15, 220
    Set txtBuscarContacto = Me.Controls.Add("Forms.TextBox.1", "txtBuscarContacto")
    Ubicar txtBuscarContacto, 235, 12, 460

    Set lstContactos = Me.Controls.Add("Forms.ListBox.1", "lstContactos")
    Ubicar lstContactos, 12, 40, 683
    With lstContactos
        .Height = 200
        .ColumnCount = 5
        .ColumnWidths = "170;220;90;200;0"       'nombre; empresa; teléfono; correo; fila (oculta)
    End With

    'Datos del contacto seleccionado
    Etiqueta "Nombre", 12, 255, 70
    Ubicar Me.Controls.Add("Forms.TextBox.1", "txtNombre"), 85, 252, 260
    Etiqueta "Correo", 365, 255, 60
    Ubicar Me.Controls.Add("Forms.TextBox.1", "txtCorreo"), 430, 252, 265
    Etiqueta "Teléfono", 12, 281, 70
    Ubicar Me.Controls.Add("Forms.TextBox.1", "txtTelefono"), 85, 278, 260
    Etiqueta "Teléfono 2", 365, 281, 60
    Ubicar Me.Controls.Add("Forms.TextBox.1", "txtTelefono2"), 430, 278, 265
    Me.Controls("txtTelefono").MaxLength = 10
    Me.Controls("txtTelefono2").MaxLength = 10
    Etiqueta "Empresa", 12, 307, 70
    Set cmbEmpresaContacto = Me.Controls.Add("Forms.ComboBox.1", "cmbEmpresaContacto")
    Ubicar cmbEmpresaContacto, 85, 304, 610
    With cmbEmpresaContacto
        .ColumnCount = 2
        .ColumnWidths = "606;0"                  'empresa; EmpresaID (oculto)
        .Style = fmStyleDropDownList
        .MatchEntry = fmMatchEntryComplete
    End With

    Set cmdGuardarContacto = Boton("cmdGuardarContacto", "Guardar cambios", 265)
    Set cmdEditarEmpresa = Boton("cmdEditarEmpresa", "Editar empresa", 375)
    Set cmdEliminarContacto = Boton("cmdEliminarContacto", "Eliminar contacto", 485)
    Set cmdCerrarContactos = Boton("cmdCerrarContactos", "Cerrar", 595)
    cmdCerrarContactos.Cancel = True

    LlenarEmpresas
    LlenarLista ""
    LimpiarContacto
End Sub

' ===============================================
' Construcción de controles
' ===============================================
Private Sub Etiqueta(ByVal texto As String, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    With Me.Controls.Add("Forms.Label.1")
        .Caption = texto
        .Left = x
        .Top = y
        .Width = w
    End With
End Sub

Private Sub Ubicar(ByVal c As Control, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    c.Left = x
    c.Top = y
    c.Width = w
End Sub

Private Function Boton(ByVal nombre As String, ByVal texto As String, _
                       ByVal x As Single) As MSForms.CommandButton
    Set Boton = Me.Controls.Add("Forms.CommandButton.1", nombre)
    With Boton
        .Caption = texto
        .Left = x
        .Top = 345
        .Width = 100
        .Height = 24
    End With
End Function

' ===============================================
' Listas
' ===============================================
Private Sub txtBuscarContacto_Change()
    LlenarLista Trim(txtBuscarContacto.Text)
End Sub

' Contactos que contienen el texto en nombre, empresa, teléfonos o correo
Private Sub LlenarLista(ByVal texto As String)
    Dim ws As Worksheet, i As Long, n As Long, empresa As String, filaEmp As Long

    Set ws = Sheets("Clientes")
    lstContactos.Clear
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If Trim(ws.Cells(i, "A").Value) <> "" Then
            'La empresa ligada manda; si no hay, la razón social escrita en el contacto
            empresa = Trim(ws.Cells(i, "B").Value)
            filaEmp = FilaEmpresaPorID(Val(ws.Cells(i, "H").Value))
            If filaEmp > 0 Then empresa = Sheets("Empresas").Cells(filaEmp, "B").Value

            If texto = "" Or _
               InStr(1, ws.Cells(i, "A").Value, texto, vbTextCompare) > 0 Or _
               InStr(1, empresa, texto, vbTextCompare) > 0 Or _
               InStr(1, ws.Cells(i, "D").Value, texto, vbTextCompare) > 0 Or _
               InStr(1, ws.Cells(i, "E").Value, texto, vbTextCompare) > 0 Or _
               InStr(1, ws.Cells(i, "F").Value, texto, vbTextCompare) > 0 Then
                lstContactos.AddItem ws.Cells(i, "A").Value
                n = lstContactos.ListCount - 1
                lstContactos.List(n, 1) = empresa
                lstContactos.List(n, 2) = ws.Cells(i, "D").Value
                lstContactos.List(n, 3) = ws.Cells(i, "F").Value
                lstContactos.List(n, 4) = i
            End If
        End If
    Next i
End Sub

' Empresas para la lista desplegable (la primera opción es "sin empresa")
Private Sub LlenarEmpresas()
    Dim ws As Worksheet, i As Long

    Set ws = Sheets("Empresas")
    cmbEmpresaContacto.Clear
    cmbEmpresaContacto.AddItem "(sin empresa)"
    cmbEmpresaContacto.List(0, 1) = 0
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        cmbEmpresaContacto.AddItem ws.Cells(i, "B").Value & "  (" & ws.Cells(i, "C").Value & ")"
        cmbEmpresaContacto.List(cmbEmpresaContacto.ListCount - 1, 1) = Val(ws.Cells(i, "A").Value)
    Next i
End Sub

Private Sub ElegirEmpresa(ByVal EmpresaID As Long)
    Dim i As Long
    For i = 0 To cmbEmpresaContacto.ListCount - 1
        If Val(cmbEmpresaContacto.List(i, 1)) = EmpresaID Then
            cmbEmpresaContacto.ListIndex = i
            Exit Sub
        End If
    Next i
    cmbEmpresaContacto.ListIndex = 0
End Sub

Private Function EmpresaElegida() As Long
    If cmbEmpresaContacto.ListIndex > 0 Then
        EmpresaElegida = Val(cmbEmpresaContacto.List(cmbEmpresaContacto.ListIndex, 1))
    End If
End Function

' ===============================================
' Contacto seleccionado
' ===============================================
Private Sub lstContactos_Click()
    If lstContactos.ListIndex = -1 Then Exit Sub
    mFila = CLng(lstContactos.List(lstContactos.ListIndex, 4))
    With Sheets("Clientes")
        Me.Controls("txtNombre").Value = .Cells(mFila, "A").Value
        Me.Controls("txtTelefono").Value = .Cells(mFila, "D").Value
        Me.Controls("txtTelefono2").Value = .Cells(mFila, "E").Value
        Me.Controls("txtCorreo").Value = .Cells(mFila, "F").Value
        ElegirEmpresa Val(.Cells(mFila, "H").Value)
    End With
End Sub

Private Sub LimpiarContacto()
    mFila = 0
    Me.Controls("txtNombre").Value = ""
    Me.Controls("txtTelefono").Value = ""
    Me.Controls("txtTelefono2").Value = ""
    Me.Controls("txtCorreo").Value = ""
    cmbEmpresaContacto.ListIndex = 0
End Sub

' ===============================================
' Botones
' ===============================================
Private Sub cmdGuardarContacto_Click()
    Dim id As Long, filaEmp As Long

    If mFila = 0 Then
        MsgBox "Selecciona un contacto de la lista.", vbExclamation, "CALEIDO"
        Exit Sub
    End If
    If Not Validar() Then Exit Sub

    id = EmpresaElegida()
    With Sheets("Clientes")
        .Cells(mFila, "A").Value = Trim(Me.Controls("txtNombre").Value)
        .Cells(mFila, "D").Value = Trim(Me.Controls("txtTelefono").Value)
        .Cells(mFila, "E").Value = Trim(Me.Controls("txtTelefono2").Value)
        .Cells(mFila, "F").Value = Trim(Me.Controls("txtCorreo").Value)
        .Cells(mFila, "H").Value = IIf(id > 0, id, Empty)
        'Copia de los datos fiscales de la empresa (si se quita la empresa, se conservan)
        filaEmp = FilaEmpresaPorID(id)
        If filaEmp > 0 Then
            .Cells(mFila, "B").Value = Sheets("Empresas").Cells(filaEmp, "B").Value
            .Cells(mFila, "C").Value = Sheets("Empresas").Cells(filaEmp, "C").Value
            .Cells(mFila, "G").Value = DomicilioEmpresa(filaEmp)
        End If
    End With

    ThisWorkbook.Save
    LlenarLista Trim(txtBuscarContacto.Text)
    MsgBox "Contacto actualizado.", vbInformation, "CALEIDO"
End Sub

Private Sub cmdEditarEmpresa_Click()
    Dim filaEmp As Long, id As Long

    id = EmpresaElegida()
    filaEmp = FilaEmpresaPorID(id)
    If filaEmp = 0 Then
        MsgBox "Elige la empresa en la lista ""Empresa"" (o selecciona un contacto ligado a una).", _
               vbExclamation, "CALEIDO"
        Exit Sub
    End If

    If frmEmpresa.Editar(filaEmp) Then
        ThisWorkbook.Save
        LlenarEmpresas
        ElegirEmpresa id
        LlenarLista Trim(txtBuscarContacto.Text)
    End If
End Sub

Private Sub cmdEliminarContacto_Click()
    Dim nombre As String

    If mFila = 0 Then
        MsgBox "Selecciona un contacto de la lista.", vbExclamation, "CALEIDO"
        Exit Sub
    End If

    nombre = Sheets("Clientes").Cells(mFila, "A").Value
    If MsgBox("¿Eliminar el contacto """ & nombre & """?" & vbCrLf & vbCrLf & _
              "Sus pedidos no se modifican. La empresa sigue registrada.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "CALEIDO") = vbNo Then Exit Sub

    Sheets("Clientes").Rows(mFila).Delete
    ThisWorkbook.Save
    LimpiarContacto
    LlenarLista Trim(txtBuscarContacto.Text)
End Sub

Private Sub cmdCerrarContactos_Click()
    Unload Me
End Sub

' Mismas reglas que el formulario de pedidos: nombre, teléfono y correo válido
Private Function Validar() As Boolean
    Dim re As Object, msg As String

    Set re = CreateObject("VBScript.RegExp")
    If Trim(Me.Controls("txtNombre").Value) = "" Then msg = msg & "- Nombre obligatorio" & vbCrLf
    re.Pattern = "^[0-9]{1,10}$"
    If Not re.Test(Trim(Me.Controls("txtTelefono").Value)) Then msg = msg & "- Teléfono: solo números (hasta 10)" & vbCrLf
    re.Pattern = "^([0-9]{1,10})?$"
    If Not re.Test(Trim(Me.Controls("txtTelefono2").Value)) Then msg = msg & "- Teléfono 2: solo números (hasta 10)" & vbCrLf
    re.Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    If Not re.Test(Trim(Me.Controls("txtCorreo").Value)) Then msg = msg & "- Correo con formato válido" & vbCrLf

    If msg <> "" Then MsgBox "Revisa:" & vbCrLf & msg, vbExclamation, "CALEIDO"
    Validar = (msg = "")
End Function
