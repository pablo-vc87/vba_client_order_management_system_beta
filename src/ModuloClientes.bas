Attribute VB_Name = "ModuloClientes"
Public Type ResultadoCliente
    Existe As Boolean
    Fila As Long
    Criterio As String ' "RFC" | "COINCIDENCIA" | ""
End Type


' ===============================================
' -------- DETECTAR CLIENTE EXISTENTE -----------
' ===============================================
Public Function DetectarClienteExistente( _
    ByVal Nombre As String, _
    ByVal RFC As String, _
    ByVal Tel1 As String, _
    ByVal Tel2 As String, _
    ByVal Email As String) As ResultadoCliente

    Dim ws As Worksheet
    Dim ultFila As Long
    Dim i As Long
    Dim res As ResultadoCliente

    Set ws = Sheets("Clientes")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' ==========================
    ' 1?? PRIORIDAD: RFC
    ' ==========================
    If Trim(RFC) <> "" Then
        For i = 2 To ultFila
            If Trim(ws.Cells(i, "C").Value) = Trim(RFC) Then
                res.Existe = True
                res.Fila = i
                res.Criterio = "RFC"
                DetectarClienteExistente = res
                Exit Function
            End If
        Next i
    End If

    ' ==========================
    ' 2?? COINCIDENCIA FLEXIBLE
    ' Nombre + (Tel o Email)
    ' ==========================
    For i = 2 To ultFila
        If UCase(Trim(ws.Cells(i, "A").Value)) = UCase(Trim(Nombre)) Then
            
            If Trim(Tel1) <> "" And Trim(ws.Cells(i, "D").Value) = Trim(Tel1) Then
                res.Existe = True
                res.Fila = i
                res.Criterio = "COINCIDENCIA"
                DetectarClienteExistente = res
                Exit Function
            End If

            If Trim(Tel2) <> "" And Trim(ws.Cells(i, "E").Value) = Trim(Tel2) Then
                res.Existe = True
                res.Fila = i
                res.Criterio = "COINCIDENCIA"
                DetectarClienteExistente = res
                Exit Function
            End If

            If Trim(Email) <> "" And Trim(ws.Cells(i, "F").Value) = Trim(Email) Then
                res.Existe = True
                res.Fila = i
                res.Criterio = "COINCIDENCIA"
                DetectarClienteExistente = res
                Exit Function
            End If

        End If
    Next i

    ' ==========================
    ' 3?? NO EXISTE
    ' ==========================
    res.Existe = False
    res.Fila = 0
    res.Criterio = ""
    DetectarClienteExistente = res

End Function

' ===============================================
'              -------- CORRECTIVO 3 --------
' -------- FUNCIÓN DECISORA CENTRAL (LA CLAVE) -----------
' ===============================================
Public Sub DecidirActualizacionCliente(ByRef Resultado As ResultadoCliente)

    Dim snap As Object

    ' CASO 1 — Cliente no existe
    If Resultado.Existe = False Then
        If MsgBox("El cliente no existe." & vbCrLf & _
                  "¿Deseas registrarlo?", vbQuestion + vbYesNo, "CALEIDO") = vbYes Then
            UserForm1.GuardarCliente
        End If
        Exit Sub
    End If

    ' Snapshot del cliente detectado
    Set snap = SnapshotClienteDesdeFila(Resultado.Fila)

    ' CASO 2 — Cotización sin RFC ? ahora con RFC
    If snap("RFC") = "" And Trim(UserForm1.txtRFC.Value) <> "" Then
        UserForm1.ActualizarCliente Resultado.Fila
        Exit Sub
    End If

    ' CASO 3 — Cambios estructurales
    If ClienteSnapshotFueModificado(snap) Then
        If MsgBox("Se modificó información del cliente." & vbCrLf & _
                  "¿Deseas actualizar la Hoja Clientes?", _
                  vbQuestion + vbYesNo, "CALEIDO") = vbYes Then
            UserForm1.ActualizarCliente Resultado.Fila
        End If
    End If

End Sub

Public Function ClienteActualDifiereDeHoja(ByVal Fila As Long) As Boolean

    With Sheets("Clientes")
        If .Cells(Fila, "A").Value <> Trim(UserForm1.txtNombreContacto.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "B").Value <> Trim(UserForm1.txtRazonSocial.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "C").Value <> Trim(UserForm1.txtRFC.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "D").Value <> Trim(UserForm1.txtTel.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "E").Value <> Trim(UserForm1.txtTel2.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "F").Value <> Trim(UserForm1.txtEmail.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
        If .Cells(Fila, "G").Value <> Trim(UserForm1.txtDomFiscal.Value) Then ClienteActualDifiereDeHoja = True: Exit Function
    End With

    ClienteActualDifiereDeHoja = False

End Function

' ===============================================
'              -------- CORRECTIVO 1 --------
' -------- CREA SNAPSHOT DE CLIENTE DETECTADO -----------
' ===============================================

Public Function SnapshotClienteDesdeFila(ByVal Fila As Long) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    With Sheets("Clientes")
        d("Nombre") = Trim(.Cells(Fila, "A").Value)
        d("Razon") = Trim(.Cells(Fila, "B").Value)
        d("RFC") = Trim(.Cells(Fila, "C").Value)
        d("Tel1") = Trim(.Cells(Fila, "D").Value)
        d("Tel2") = Trim(.Cells(Fila, "E").Value)
        d("Email") = Trim(.Cells(Fila, "F").Value)
        d("DomFiscal") = Trim(.Cells(Fila, "G").Value)
    End With

    Set SnapshotClienteDesdeFila = d

End Function

' ===============================================
'              -------- CORRECTIVO 2 --------
' -------- COMPARA SOLO CAMPOS ESTRUCTURALES -----------
' ===============================================
Public Function ClienteSnapshotFueModificado(ByVal snap As Object) As Boolean

    If snap Is Nothing Then Exit Function

    If snap("Nombre") <> Trim(UserForm1.txtNombreContacto.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("Razon") <> Trim(UserForm1.txtRazonSocial.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("RFC") <> Trim(UserForm1.txtRFC.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("Tel1") <> Trim(UserForm1.txtTel.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("Tel2") <> Trim(UserForm1.txtTel2.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("Email") <> Trim(UserForm1.txtEmail.Value) Then ClienteSnapshotFueModificado = True: Exit Function
    If snap("DomFiscal") <> Trim(UserForm1.txtDomFiscal.Value) Then ClienteSnapshotFueModificado = True: Exit Function

    ClienteSnapshotFueModificado = False

End Function
