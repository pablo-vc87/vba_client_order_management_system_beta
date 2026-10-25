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
' VALIDAR N⁄MERO POSITIVO
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
' VALIDAR CORREO ELECTR”NICO (ESTRUCTURA)
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

'FUNCI”N QUE VALIDA EL CORREO ELECTR”NICO
Function Valida_Correo(Email As String) As Boolean
    Application.Volatile
    'Declaramos variables
    Dim oReg As RegExp
    'Crea un Nuevo objeto RegExp
    Set oReg = New RegExp
    On Error GoTo ErrorHandler
    'ExpresiÛn regular para validar direcciones .com
    'oReg.Pattern = "^[\w-\.]+@\w+\.\w+$"
    ' ExpresiÛn regular para validar direcciones .com.pe
    oReg.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    ' Comprueba y Retorna TRue o false
    Valida_Correo = oReg.Test(Email)
    Set oReg = Nothing
    Exit Function
    'En caso de error
ErrorHandler:
    MsgBox "Ha ocurrido un error: ", vbExclamation, "El TÌo Tech"
End Function

'=================================================
' VALIDAR P¡GINA 1 (CLIENTE)
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
' VALIDAR P¡GINA 2 (PRODUCTOS DIN¡MICOS)
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
' VALIDAR FECHAS
' Entrega y recepciones no pueden ser anteriores a la fecha del pedido,
' y ninguna recepciÛn puede ser posterior a la entrega (el mismo dÌa sÌ se permite)
'=================================================
Public Function ValidarFechas() As Boolean

    Dim ctrl As Control
    Dim fPedido As Date, fEntrega As Date
    Dim hayEntrega As Boolean
    Dim ok As Boolean
    Dim msg As String

    ok = True

    With UserForm1
        If Not IsDate(.txtFechaActual.Value) Then
            ValidarFechas = True
            Exit Function
        End If
        fPedido = CDate(.txtFechaActual.Value)

        'FECHA DE ENTREGA
        .txtFechaEntrega.BackColor = vbWhite
        If Trim(.txtFechaEntrega.Value) <> "" Then
            hayEntrega = True
            fEntrega = CDate(.txtFechaEntrega.Value)
            If fEntrega < fPedido Then
                .txtFechaEntrega.BackColor = RGB(255, 180, 180)
                msg = msg & "- La fecha de entrega es anterior a la fecha del pedido." & vbCrLf
                ok = False
            End If
        End If

        'FECHAS DE RECEPCI”N DE CADA PRODUCTO
        For Each ctrl In .frProductos.Controls
            If ctrl.Name Like "txtFechaRec*" Then
                ctrl.BackColor = RGB(240, 240, 240)
                If Trim(ctrl.Value) <> "" Then
                    If CDate(ctrl.Value) < fPedido Then
                        ctrl.BackColor = RGB(255, 180, 180)
                        msg = msg & "- Producto " & Mid(ctrl.Name, 12) & _
                              ": la recepciÛn es anterior a la fecha del pedido." & vbCrLf
                        ok = False
                    ElseIf hayEntrega Then
                        If CDate(ctrl.Value) > fEntrega Then
                            ctrl.BackColor = RGB(255, 180, 180)
                            msg = msg & "- Producto " & Mid(ctrl.Name, 12) & _
                                  ": la recepciÛn es posterior a la entrega." & vbCrLf
                            ok = False
                        End If
                    End If
                End If
            End If
        Next ctrl
    End With

    If Not ok Then
        MsgBox "Revise las fechas:" & vbCrLf & vbCrLf & msg, vbExclamation, "CALEIDO"
    End If

    ValidarFechas = ok

End Function

'=================================================
' FECHA DEL FORMULARIO -> VALOR DE CELDA
' Si se escribe el texto "05/02/2026" en una celda, Excel lo lee como mes/dÌa
' (5 de mayo). CDate respeta la configuraciÛn regional (dÌa/mes).
'=================================================
Public Function ValorFecha(ByVal v As Variant) As Variant
    If Trim(CStr(v)) = "" Then
        ValorFecha = Empty
    ElseIf IsDate(v) Then
        ValorFecha = CDate(v)
    Else
        ValorFecha = v
    End If
End Function

'=================================================
' N⁄MERO DE CELDA -> TEXTO DEL FORMULARIO
' Asignar un n˙mero a un TextBox usa el separador decimal de Windows (coma), pero
' los campos solo aceptan punto y Val("1500,5") da 1500. Str siempre usa punto.
'=================================================
Public Function NumeroATexto(ByVal v As Variant) As String
    If IsNumeric(v) And Trim(CStr(v)) <> "" Then
        NumeroATexto = Trim(Str(CDbl(v)))
        If Left(NumeroATexto, 1) = "." Then NumeroATexto = "0" & NumeroATexto
    Else
        NumeroATexto = CStr(v)
    End If
End Function

'=================================================
' REPARAR FECHAS GUARDADAS COMO TEXTO (correr una vez)
' Pedidos B y K, Detalle_Pedidos E
'=================================================
Public Sub RepararFechasTexto()
    Dim c As Range, n As Long

    For Each c In Union(Sheets("Pedidos").Range("B2:B" & UltimaFila("Pedidos")), _
                        Sheets("Pedidos").Range("K2:K" & UltimaFila("Pedidos")))
        If ConvertirCeldaFecha(c) Then n = n + 1
    Next c
    For Each c In Sheets("Detalle_Pedidos").Range("E2:E" & UltimaFila("Detalle_Pedidos"))
        If ConvertirCeldaFecha(c) Then n = n + 1
    Next c

    Sheets("Pedidos").Range("B:B,K:K").NumberFormat = "dd/mm/yyyy"
    Sheets("Detalle_Pedidos").Range("E:E").NumberFormat = "dd/mm/yyyy"
    MsgBox n & " fechas convertidas de texto a fecha.", vbInformation, "CALEIDO"
End Sub

Private Function ConvertirCeldaFecha(ByVal c As Range) As Boolean
    If VarType(c.Value) = vbString Then
        If IsDate(c.Value) Then
            c.Value = CDate(c.Value)
            ConvertirCeldaFecha = True
        End If
    End If
End Function

Private Function UltimaFila(ByVal hoja As String) As Long
    UltimaFila = Sheets(hoja).Cells(Sheets(hoja).Rows.Count, "A").End(xlUp).Row
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
        If c Like "[A-Za-z¡…Õ”⁄·ÈÌÛ˙—Ò ]" Then
            Resultado = Resultado & c
        End If
    Next i

    LimpiarSoloTexto = Resultado
End Function


