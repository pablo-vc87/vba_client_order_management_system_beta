Attribute VB_Name = "ModuloValidaciones"
Option Explicit

'=================================================
' FUNCIONES BASE
'=================================================
'=================================================
' VALIDAR SOLO TEXTO
'=================================================
Private Function Requerido(ctrl As Control) As Boolean
    If Trim(ctrl.Value) = "" Then
        ctrl.BackColor = RGB(243, 233, 25) ' amarillo
        Requerido = False
    Else
        ctrl.BackColor = vbWhite
        Requerido = True
    End If
End Function

'=================================================
' VALIDAR NÚMERO POSITIVO
'=================================================
Private Function NumeroPositivo(ctrl As Control) As Boolean

    If Not IsNumeric(ctrl.Value) Or Val(ctrl.Value) <= 0 Then
        ctrl.BackColor = RGB(255, 180, 180) ' rojo claro
        NumeroPositivo = False
    Else
        ctrl.BackColor = vbWhite
        NumeroPositivo = True
    End If

End Function

'=================================================
' VALIDAR CORREO ELECTRÓNICO (ESTRUCTURA)
'=================================================
Private Function CorreoValido(tb As MSForms.TextBox) As Boolean

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    re.IgnoreCase = True

    If Trim(tb.Value) = "" Or Not re.Test(tb.Value) Then
        tb.BackColor = RGB(255, 180, 180)
        CorreoValido = False
    Else
        tb.BackColor = vbWhite
        CorreoValido = True
    End If

End Function

'FUNCIÓN QUE VALIDA EL CORREO ELECTRÓNICO
Function Valida_Correo(Email As String) As Boolean
    Application.Volatile
    'Declaramos variables
    Dim oReg As RegExp
    'Crea un Nuevo objeto RegExp
    Set oReg = New RegExp
    On Error GoTo ErrorHandler
    'Expresión regular para validar direcciones .com
    'oReg.Pattern = "^[\w-\.]+@\w+\.\w+$"
    ' Expresión regular para validar direcciones .com.pe
    oReg.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    ' Comprueba y Retorna TRue o false
    Valida_Correo = oReg.Test(Email)
    Set oReg = Nothing
    Exit Function
    'En caso de error
ErrorHandler:
    MsgBox "Ha ocurrido un error: ", vbExclamation, "El Tío Tech"
End Function

'=================================================
' VALIDAR PÁGINA 1 (CLIENTE)
'=================================================
Public Function ValidarPagina1() As Boolean

    Dim ok As Boolean
    ok = True

    With UserForm1
        ok = Requerido(.txtNombreContacto) And ok
        ok = CorreoValido(.txtEmail) And ok
        ok = Requerido(.txtTel) And ok
        ok = Requerido(.cmbEstatus) And ok
    End With

    If Not ok Then
        MsgBox "Revise los datos del cliente", vbExclamation, "CALEIDO"
    End If

    ValidarPagina1 = ok

End Function

'=================================================
' VALIDAR PÁGINA 2 (PRODUCTOS DINÁMICOS)
'=================================================
Public Function ValidarPagina2() As Boolean

    Dim ctrl As Control
    Dim ok As Boolean
    Dim hayProducto As Boolean

    ok = True
    hayProducto = False

    With UserForm1.frProductos

        For Each ctrl In .Controls

            If ctrl.Name Like "cmbTec*" Then hayProducto = True

            Select Case True
                Case ctrl.Name Like "cmbTec*"
                    ok = Requerido(ctrl) And ok

                Case ctrl.Name Like "txtCant*"
                    ok = NumeroPositivo(ctrl) And ok

                Case ctrl.Name Like "txtPrecio*"
                    ok = NumeroPositivo(ctrl) And ok

                Case ctrl.Name Like "txtLogo*"
                    ok = Requerido(ctrl) And ok

                Case ctrl.Name Like "txtTam*"
                    ok = Requerido(ctrl) And ok
            End Select

        Next ctrl

    End With

    If Not hayProducto Then
        MsgBox "Debe agregar al menos un producto", vbExclamation, "CALEIDO"
        ValidarPagina2 = False
        Exit Function
    End If

    If Not ok Then
        MsgBox "Revise los datos de los productos", vbExclamation, "CALEIDO"
    End If

    ValidarPagina2 = ok

End Function

'=================================================
' LIMPIAR TEXTO (SOLO LETRAS)
'=================================================
Public Function LimpiarSoloTexto(ByVal texto As String) As String
    Dim i As Long
    Dim c As String
    Dim Resultado As String

    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        If c Like "[A-Za-zÁÉÍÓÚáéíóúÑñ ]" Then
            Resultado = Resultado & c
        End If
    Next i

    LimpiarSoloTexto = Resultado
End Function


