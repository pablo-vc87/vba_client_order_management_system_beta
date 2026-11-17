VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFacturar 
   Caption         =   "CALEIDO | Facturación"
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Datos que se piden al generar la prefactura (forma, método y uso del CFDI) o al
' registrar un pago (fecha, monto, forma y UUID). Los controles se crean al llamar a
' PedirDatosFactura o PedirPago: no hay que dibujarlos.

Private mOK As Boolean
Private mModo As String
Private mSaldo As Double
Private mTop As Single
Private WithEvents cmbMetodo As MSForms.ComboBox
Private WithEvents cmdAceptar As MSForms.CommandButton
Private WithEvents cmdCancelar As MSForms.CommandButton

' ===============================================
' PREFACTURA: devuelve True si se aceptó; forma, metodo y uso salen con la clave elegida
' ===============================================
Public Function PedirDatosFactura(ByRef forma As String, ByRef metodo As String, _
                                  ByRef uso As String, ByVal total As Double) As Boolean
    mModo = "FACTURA"
    Me.Caption = "CALEIDO | Prefactura"
    Aviso "Total del pedido: $" & Format(total, "#,##0.00")

    Set cmbMetodo = Lista("Metodo", "Método de pago", ListaMetodosPago())
    Lista "Forma", "Forma de pago", ListaFormasPago(True)
    Lista "Uso", "Uso del CFDI", ListaUsosCFDI()
    Aviso "PUE: se paga todo al facturar.  PPD: anticipo o parcialidades (forma 99, " & _
          "y un complemento de pago por cada abono)."
    Botones

    Elegir "Metodo", IIf(metodo = "", "PUE", metodo)
    Elegir "Forma", IIf(forma = "", IIf(metodo = "PPD", "99", "03"), forma)
    Elegir "Uso", IIf(uso = "", "G03", uso)

    Me.Show
    If mOK Then
        metodo = ClaveCatalogo(Ctl("Metodo").Value)
        forma = ClaveCatalogo(Ctl("Forma").Value)
        uso = ClaveCatalogo(Ctl("Uso").Value)
    End If
    PedirDatosFactura = mOK
    Unload Me
End Function

' ===============================================
' PAGO: devuelve True si se aceptó; fecha, monto, forma y uuid salen con lo capturado
' ===============================================
Public Function PedirPago(ByVal total As Double, ByVal pagado As Double, ByVal saldo As Double, _
                          ByRef fecha As Date, ByRef monto As Double, ByRef forma As String, _
                          ByRef uuid As String) As Boolean
    mModo = "PAGO"
    mSaldo = saldo
    Me.Caption = "CALEIDO | Registrar pago"
    Aviso "Total: $" & Format(total, "#,##0.00") & "    Pagado: $" & Format(pagado, "#,##0.00") & _
          "    Saldo: $" & Format(saldo, "#,##0.00")

    Texto "Fecha", "Fecha de pago (dd/mm/aaaa)", Format(fecha, "dd/mm/yyyy")
    Texto "Monto", "Monto pagado", Trim(Str(Round(monto, 2)))
    Lista "Forma", "Forma de pago", ListaFormasPago(False)
    Texto "UUID", "UUID de la factura timbrada", uuid
    Aviso "El UUID está en el XML o PDF de la factura timbrada (folio fiscal). " & _
          "Se guarda en Pedidos para los siguientes pagos."
    Botones

    Elegir "Forma", IIf(forma = "", "03", forma)

    Me.Show
    If mOK Then
        fecha = CDate(Ctl("Fecha").Value)
        monto = LeerMonto(Ctl("Monto").Value)
        forma = ClaveCatalogo(Ctl("Forma").Value)
        uuid = UCase(Trim(Ctl("UUID").Value))
    End If
    PedirPago = mOK
    Unload Me
End Function

' ===============================================
' Construcción de controles
' ===============================================
Private Sub UserForm_Initialize()
    Me.Width = 470
    mTop = 12
End Sub

Private Sub Aviso(ByVal texto As String)
    With Me.Controls.Add("Forms.Label.1")
        .Caption = texto
        .Left = 12
        .Top = mTop
        .Width = 430
        .Height = 26
        .WordWrap = True
    End With
    mTop = mTop + 30
End Sub

Private Function Lista(ByVal nombre As String, ByVal etiqueta As String, _
                       ByVal elementos As Variant) As MSForms.ComboBox
    Dim e As Variant
    Etiquetar etiqueta
    Set Lista = Me.Controls.Add("Forms.ComboBox.1", "ctl" & nombre)
    With Lista
        .Left = 170
        .Top = mTop
        .Width = 272
        .Style = fmStyleDropDownList
        For Each e In elementos
            .AddItem e
        Next e
    End With
    mTop = mTop + 26
End Function

Private Sub Texto(ByVal nombre As String, ByVal etiqueta As String, ByVal valor As String)
    Etiquetar etiqueta
    With Me.Controls.Add("Forms.TextBox.1", "ctl" & nombre)
        .Left = 170
        .Top = mTop
        .Width = 272
        .Value = valor
    End With
    mTop = mTop + 26
End Sub

Private Sub Etiquetar(ByVal etiqueta As String)
    With Me.Controls.Add("Forms.Label.1")
        .Caption = etiqueta
        .Left = 12
        .Top = mTop + 3
        .Width = 155
    End With
End Sub

Private Sub Botones()
    mTop = mTop + 6
    Set cmdAceptar = Me.Controls.Add("Forms.CommandButton.1", "cmdAceptar")
    With cmdAceptar
        .Caption = "Aceptar"
        .Left = 232
        .Top = mTop
        .Width = 100
        .Height = 24
        .Default = True
    End With
    Set cmdCancelar = Me.Controls.Add("Forms.CommandButton.1", "cmdCancelar")
    With cmdCancelar
        .Caption = "Cancelar"
        .Left = 342
        .Top = mTop
        .Width = 100
        .Height = 24
        .Cancel = True
    End With
    Me.Height = mTop + 24 + 40
End Sub

Private Function Ctl(ByVal nombre As String) As Control
    Set Ctl = Me.Controls("ctl" & nombre)
End Function

' Selecciona en una lista el elemento con esa clave
Private Sub Elegir(ByVal nombre As String, ByVal clave As String)
    Dim i As Long
    With Ctl(nombre)
        For i = 0 To .ListCount - 1
            If ClaveCatalogo(.List(i)) = clave Then
                .ListIndex = i
                Exit Sub
            End If
        Next i
    End With
End Sub

' PPD siempre va con forma 99 (por definir); PUE nunca
Private Sub cmbMetodo_Change()
    If ClaveCatalogo(cmbMetodo.Value) = "PPD" Then
        Elegir "Forma", "99"
    ElseIf ClaveCatalogo(Ctl("Forma").Value) = "99" Then
        Elegir "Forma", "03"
    End If
End Sub

' ===============================================
' Aceptar / cancelar
' ===============================================
Private Sub cmdAceptar_Click()
    If Not Validar() Then Exit Sub
    mOK = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        mOK = False
        Me.Hide
    End If
End Sub

' "1500.50", "1,500.50" o "1500,50" -> 1500.5 (la coma es decimal solo si no hay punto)
Private Function LeerMonto(ByVal s As String) As Double
    s = Trim(s)
    If InStr(s, ".") = 0 Then s = Replace(s, ",", ".") Else s = Replace(s, ",", "")
    LeerMonto = Round(Val(s), 2)
End Function

Private Function Validar() As Boolean
    Dim m As Double, re As Object

    If mModo = "FACTURA" Then
        If Ctl("Metodo").ListIndex = -1 Or Ctl("Forma").ListIndex = -1 Or Ctl("Uso").ListIndex = -1 Then
            MsgBox "Elige método, forma de pago y uso del CFDI.", vbExclamation, "CALEIDO"
            Exit Function
        End If
        If (ClaveCatalogo(Ctl("Metodo").Value) = "PPD") <> (ClaveCatalogo(Ctl("Forma").Value) = "99") Then
            MsgBox "Con PPD la forma de pago es 99 (por definir); con PUE no puede ser 99.", _
                   vbExclamation, "CALEIDO"
            Exit Function
        End If
    Else
        If Not IsDate(Ctl("Fecha").Value) Then
            MsgBox "Escribe la fecha de pago como dd/mm/aaaa.", vbExclamation, "CALEIDO"
            Exit Function
        End If
        If CDate(Ctl("Fecha").Value) > Date Then
            MsgBox "La fecha de pago no puede ser futura.", vbExclamation, "CALEIDO"
            Exit Function
        End If
        m = LeerMonto(Ctl("Monto").Value)
        If m <= 0 Or m > Round(mSaldo, 2) Then
            MsgBox "El monto debe ser mayor que 0 y no mayor que el saldo ($" & _
                   Format(mSaldo, "#,##0.00") & ").", vbExclamation, "CALEIDO"
            Exit Function
        End If
        If Ctl("Forma").ListIndex = -1 Then
            MsgBox "Elige la forma de pago.", vbExclamation, "CALEIDO"
            Exit Function
        End If
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})?$"
        If Not re.Test(Trim(Ctl("UUID").Value)) Then
            MsgBox "El UUID tiene el formato 8-4-4-4-12 (ej. 6F9619FF-8B86-D011-B42D-00C04FC964FF)." & _
                   vbCrLf & "Déjalo vacío si la factura aún no se timbra.", vbExclamation, "CALEIDO"
            Exit Function
        End If
    End If

    Validar = True
End Function
