VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmpresa 
   Caption         =   "CALEIDO | Registrar empresa"
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Alta de una empresa nueva en la hoja Empresas.
' Los controles se crean al abrir el formulario: no hay que dibujarlos.
' Uso: EmpresaID = frmEmpresa.Registrar(razon, rfc, email, domicilio)   (0 si se cancela)

' Campo|Etiqueta, en el orden de las columnas B a N de la hoja Empresas
Private Const CAMPOS As String = "RazonSocial|Razón social *;RFC|RFC *;Calle|Calle;" & _
    "NumExterior|Núm. exterior;NumInterior|Núm. interior;Colonia|Colonia;CP|C.P.;" & _
    "Municipio|Municipio / alcaldía;Estado|Estado;Pais|País;RegimenFiscal|Régimen fiscal;" & _
    "UsoCFDI|Uso de CFDI;EmailFacturacion|Correo de facturación"

Private mID As Long
Private lblDomicilio As MSForms.Label
Private WithEvents cmdGuardarEmp As MSForms.CommandButton
Private WithEvents cmdCancelarEmp As MSForms.CommandButton

Public Function Registrar(ByVal razon As String, ByVal RFC As String, _
                          ByVal Email As String, ByVal domicilio As String) As Long
    mID = 0
    Ctl("RazonSocial").Value = Trim(razon)
    Ctl("RFC").Value = UCase(Trim(RFC))
    Ctl("EmailFacturacion").Value = Trim(Email)
    If Trim(domicilio) <> "" Then
        lblDomicilio.Caption = "Domicilio escrito en el pedido: " & Trim(domicilio)
    End If

    Me.Show                         'modal: sigue aquí cuando se guarda o se cancela
    Registrar = mID
    Unload Me
End Function

Private Sub UserForm_Initialize()
    Dim lista() As String, p() As String, i As Long, t As Single

    Me.Caption = "CALEIDO | Registrar empresa"
    Me.Width = 540
    Me.Height = 430

    lista = Split(CAMPOS, ";")
    For i = 0 To UBound(lista)
        p = Split(lista(i), "|")
        t = 12 + i * 24

        With Me.Controls.Add("Forms.Label.1", "lbl" & p(0))
            .Caption = p(1)
            .Left = 12
            .Top = t + 3
            .Width = 125
        End With

        If p(0) = "RegimenFiscal" Or p(0) = "UsoCFDI" Then
            With Me.Controls.Add("Forms.ComboBox.1", "txt" & p(0))
                .Left = 140
                .Top = t
                .Width = 370
                .Style = fmStyleDropDownList
            End With
        Else
            With Me.Controls.Add("Forms.TextBox.1", "txt" & p(0))
                .Left = 140
                .Top = t
                .Width = 370
            End With
        End If
    Next i

    'Catálogos del SAT (se guarda solo la clave)
    With Ctl("RegimenFiscal")
        .AddItem "601 - General de Ley Personas Morales"
        .AddItem "603 - Personas Morales con Fines no Lucrativos"
        .AddItem "605 - Sueldos y Salarios"
        .AddItem "606 - Arrendamiento"
        .AddItem "610 - Residentes en el Extranjero"
        .AddItem "612 - Personas Físicas con Actividades Empresariales y Profesionales"
        .AddItem "616 - Sin obligaciones fiscales"
        .AddItem "621 - Incorporación Fiscal"
        .AddItem "623 - Opcional para Grupos de Sociedades"
        .AddItem "624 - Coordinados"
        .AddItem "625 - Actividades Empresariales con ingresos a través de Plataformas Tecnológicas"
        .AddItem "626 - Régimen Simplificado de Confianza"
    End With
    With Ctl("UsoCFDI")
        .AddItem "G01 - Adquisición de mercancías"
        .AddItem "G02 - Devoluciones, descuentos o bonificaciones"
        .AddItem "G03 - Gastos en general"
        .AddItem "S01 - Sin efectos fiscales"
        .AddItem "CP01 - Pagos"
        .Value = "G03 - Gastos en general"
    End With
    Ctl("Pais").Value = "MEXICO"
    Ctl("CP").MaxLength = 5
    Ctl("RFC").MaxLength = 13

    Set lblDomicilio = Me.Controls.Add("Forms.Label.1", "lblDomicilio")
    With lblDomicilio
        .Left = 12
        .Top = 12 + (UBound(lista) + 1) * 24
        .Width = 500
        .Height = 24
        .ForeColor = RGB(90, 90, 90)
        .WordWrap = True
    End With

    Set cmdGuardarEmp = Me.Controls.Add("Forms.CommandButton.1", "cmdGuardarEmp")
    With cmdGuardarEmp
        .Caption = "Guardar empresa"
        .Left = 300
        .Top = lblDomicilio.Top + 30
        .Width = 100
        .Height = 24
    End With
    Set cmdCancelarEmp = Me.Controls.Add("Forms.CommandButton.1", "cmdCancelarEmp")
    With cmdCancelarEmp
        .Caption = "Cancelar"
        .Left = 410
        .Top = lblDomicilio.Top + 30
        .Width = 100
        .Height = 24
        .Cancel = True
    End With
End Sub

Private Function Ctl(ByVal campo As String) As Control
    Set Ctl = Me.Controls("txt" & campo)
End Function

'=================================================
' GUARDAR: valida y agrega la fila a la hoja Empresas
'=================================================
Private Sub cmdGuardarEmp_Click()
    Dim ws As Worksheet, Fila As Long, i As Long
    Dim lista() As String, campo As String, v As String

    If Not Validar() Then Exit Sub

    If BuscarEmpresa(Ctl("RFC").Value, Ctl("RazonSocial").Value) > 0 Then
        MsgBox "Ya existe una empresa con ese RFC (o con esa razón social, si el RFC es genérico).", _
               vbExclamation, "CALEIDO"
        Exit Sub
    End If

    Set ws = Sheets("Empresas")
    Fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    mID = NuevoEmpresaID()
    ws.Cells(Fila, "A").Value = mID

    lista = Split(CAMPOS, ";")
    For i = 0 To UBound(lista)
        campo = Split(lista(i), "|")(0)
        v = Trim(Ctl(campo).Value)
        With ws.Cells(Fila, i + 2)
            Select Case campo
                Case "RFC":           .Value = UCase(v)
                Case "CP":            .NumberFormat = "@": .Value = v
                Case "RegimenFiscal": If v <> "" Then .Value = Val(v)            '"601 - ..." -> 601
                Case "UsoCFDI":       If v <> "" Then .Value = Split(v, " ")(0)  '"G03 - ..." -> G03
                Case Else:            .Value = v
            End Select
        End With
    Next i

    Me.Hide
End Sub

Private Sub cmdCancelarEmp_Click()
    mID = 0
    Me.Hide
End Sub

'Cerrar con la X equivale a cancelar
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mID = 0
        Me.Hide
    End If
End Sub

'=================================================
' VALIDACIONES
'=================================================
Private Function Validar() As Boolean
    Dim ok As Boolean
    ok = True
    ok = Marcar(Ctl("RazonSocial"), Trim(Ctl("RazonSocial").Value) <> "") And ok
    ok = Marcar(Ctl("RFC"), Cumple(UCase(Trim(Ctl("RFC").Value)), "^[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}$")) And ok
    ok = Marcar(Ctl("CP"), Cumple(Trim(Ctl("CP").Value), "^([0-9]{5})?$")) And ok
    ok = Marcar(Ctl("EmailFacturacion"), _
                Cumple(Trim(Ctl("EmailFacturacion").Value), "^([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})?$")) And ok
    If Not ok Then
        MsgBox "Revisa los campos marcados:" & vbCrLf & _
               "- Razón social obligatoria" & vbCrLf & _
               "- RFC de 12 (empresa) o 13 (persona física) caracteres" & vbCrLf & _
               "- C.P. de 5 dígitos" & vbCrLf & _
               "- Correo con formato válido", vbExclamation, "CALEIDO"
    End If
    Validar = ok
End Function

Private Function Marcar(ByVal c As Control, ByVal ok As Boolean) As Boolean
    If ok Then c.BackColor = vbWhite Else c.BackColor = RGB(255, 180, 180)
    Marcar = ok
End Function

Private Function Cumple(ByVal texto As String, ByVal patron As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = patron
    Cumple = re.Test(texto)
End Function
