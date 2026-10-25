Attribute VB_Name = "ModuloEmpresas"
Option Explicit

Public Function SQL_ImportarEmpresas() As String
    SQL_ImportarEmpresas = _
        "SELECT FIRST(RazonSocial) AS RazonSocial" & _
        ", FIRST(RFC) AS RFC" & _
        ", FIRST(Calle) AS Calle" & _
        ", FIRST(NumExterior) AS NumExterior" & _
        ", FIRST(NumInterior) AS NumInterior" & _
        ", FIRST(Colonia) AS Colonia" & _
        ", FIRST(IIF(CP IS NULL, Null, Format(CP, '00000'))) AS CP" & _
        ", FIRST(Municipio) AS Municipio" & _
        ", FIRST(Estado) AS Estado" & _
        ", FIRST(Pais) AS Pais" & _
        ", FIRST(RegimenFiscal) AS RegimenFiscal" & _
        ", FIRST(CfdiUse) AS UsoCFDI" & _
        ", MAX(IIF(Email = '?', Null, Email)) AS EmailFacturacion" & _
        ", FIRST(NumRegIdTrib) AS NumRegIdTrib" & _
        ", FIRST(TaxResidence) AS ResidenciaFiscal" & _
        " FROM " & TablaLista() & _
        " GROUP BY RFC, IIF(RFC LIKE 'X%010101000', RazonSocial, '')" & _
        " ORDER BY FIRST(RazonSocial)"
End Function


Public Sub ImportarEmpresas()
    Dim ws As Worksheet, wsImp As Worksheet, n As Long, i As Long, c As Range

    ThisWorkbook.Save
    VolcarResultado SQL_ImportarEmpresas(), "Import"

    Set wsImp = ThisWorkbook.Sheets("Import")
    n = wsImp.Cells(wsImp.Rows.Count, "A").End(xlUp).Row - 1
    If n < 1 Then
        MsgBox "La consulta no devolvió empresas", vbExclamation
        Exit Sub
    End If

    If MsgBox("Se borrarán las empresas actuales y se cargarán " & n & "." & vbCrLf & _
              "Los EmpresaID se vuelven a numerar: los contactos y pedidos ya ligados " & _
              "quedarían apuntando a otra empresa. ¿Continuar?", _
              vbQuestion + vbYesNo, "CALEIDO") = vbNo Then Exit Sub

    Set ws = ObtenerHoja("Empresas")
    ws.Cells.Clear

    ws.Range("A1").Value = "EmpresaID"
    ws.Range("B1").Resize(1, 15).Value = wsImp.Range("A1").Resize(1, 15).Value
    ws.Columns("H").NumberFormat = "@"                 'CP como texto
    ws.Range("B2").Resize(n, 15).Value = wsImp.Range("A2").Resize(n, 15).Value

    For i = 1 To n
        ws.Cells(i + 1, "A").Value = i
    Next i
    For Each c In ws.Range("H2").Resize(n, 1)
        If c.Value <> "" Then c.Value = Format(c.Value, "00000")
    Next c

    ThisWorkbook.Sheets("Clientes").Range("H1").Value = "EmpresaID"
    wsImp.Cells.Clear
    ThisWorkbook.Save
End Sub

' ===============================================
' EMPRESAS (datos fiscales). Columnas de la hoja Empresas:
' A EmpresaID, B RazonSocial, C RFC, D Calle, E NumExterior, F NumInterior, G Colonia,
' H CP, I Municipio, J Estado, K Pais, L RegimenFiscal, M UsoCFDI, N EmailFacturacion,
' O NumRegIdTrib, P ResidenciaFiscal
' Clientes!H y Pedidos!P guardan el EmpresaID del contacto y del pedido.
' ===============================================

' RFC genérico (público en general o extranjero): no identifica a una sola empresa
Public Function EsRFCGenerico(ByVal RFC As String) As Boolean
    RFC = UCase(Trim(RFC))
    EsRFCGenerico = (RFC = "XAXX010101000" Or RFC = "XEXX010101000")
End Function

' Fila de la empresa en la hoja Empresas (0 si no existe). Busca por RFC; solo si el RFC
' está vacío o es genérico, busca por razón social.
Public Function BuscarEmpresa(ByVal RFC As String, ByVal razon As String) As Long
    Dim ws As Worksheet, i As Long, ult As Long

    Set ws = Sheets("Empresas")
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    RFC = UCase(Trim(RFC))
    razon = UCase(Trim(razon))

    If RFC <> "" And Not EsRFCGenerico(RFC) Then
        For i = 2 To ult
            If UCase(Trim(ws.Cells(i, "C").Value)) = RFC Then
                BuscarEmpresa = i
                Exit Function
            End If
        Next i
    ElseIf razon <> "" Then
        For i = 2 To ult
            If UCase(Trim(ws.Cells(i, "B").Value)) = razon Then
                BuscarEmpresa = i
                Exit Function
            End If
        Next i
    End If
End Function

' Fila de la empresa con ese EmpresaID (0 si no existe)
Public Function FilaEmpresaPorID(ByVal EmpresaID As Long) As Long
    Dim ws As Worksheet, i As Long, ult As Long

    If EmpresaID <= 0 Then Exit Function
    Set ws = Sheets("Empresas")
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ult
        If Val(ws.Cells(i, "A").Value) = EmpresaID Then
            FilaEmpresaPorID = i
            Exit Function
        End If
    Next i
End Function

Public Function EmpresaIDDeFila(ByVal Fila As Long) As Long
    EmpresaIDDeFila = Val(Sheets("Empresas").Cells(Fila, "A").Value)
End Function

Public Function NuevoEmpresaID() As Long
    NuevoEmpresaID = Application.WorksheetFunction.Max(Sheets("Empresas").Columns("A")) + 1
End Function

' Domicilio fiscal en una línea (mismo formato que ExprDomicilio en ModuloSQL)
Public Function DomicilioEmpresa(ByVal Fila As Long) As String
    Dim s As String
    With Sheets("Empresas")
        s = ParteDom(.Cells(Fila, "D").Value, ", ") & ParteDom(.Cells(Fila, "E").Value, " ") & _
            ParteDom(.Cells(Fila, "F").Value, " Int. ") & ParteDom(.Cells(Fila, "G").Value, ", Col. ") & _
            ParteDom(.Cells(Fila, "H").Value, ", C.P. ") & ParteDom(.Cells(Fila, "I").Value, ", ") & _
            ParteDom(.Cells(Fila, "J").Value, ", ")
    End With
    Do While Left(s, 1) = "," Or Left(s, 1) = " "
        s = Mid(s, 2)
    Loop
    DomicilioEmpresa = s
End Function

Private Function ParteDom(ByVal v As Variant, ByVal prefijo As String) As String
    If Trim(CStr(v)) <> "" Then ParteDom = prefijo & Trim(CStr(v))
End Function

' Llena la lista del formulario con las empresas cuya razón social o RFC contienen el texto.
' Columna 0: lo que se ve; columna 1 (oculta): fila en la hoja Empresas.
Public Sub LlenarListaEmpresas(ByVal lst As MSForms.ListBox, ByVal texto As String)
    Dim ws As Worksheet, i As Long, ult As Long

    Set ws = Sheets("Empresas")
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lst.Clear
    For i = 2 To ult
        If InStr(1, ws.Cells(i, "B").Value, texto, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, "C").Value, texto, vbTextCompare) > 0 Then
            lst.AddItem ws.Cells(i, "B").Value & "  (" & ws.Cells(i, "C").Value & ")"
            lst.List(lst.ListCount - 1, 1) = i
        End If
    Next i
End Sub

' Encabezados de las columnas de enlace (por si el libro todavía no los tiene)
Private Sub AsegurarEncabezados()
    If Sheets("Pedidos").Range("P1").Value = "" Then Sheets("Pedidos").Range("P1").Value = "EmpresaID"
    If Sheets("Clientes").Range("H1").Value = "" Then Sheets("Clientes").Range("H1").Value = "EmpresaID"
End Sub

' ===============================================
' EmpresaID del pedido en captura (0 = sin empresa, ej. cotización sin datos fiscales).
' Si la empresa no existe, ofrece registrarla con frmEmpresa y pasa sus datos al formulario.
' ===============================================
Public Function ResolverEmpresa() As Long
    Dim Fila As Long, razon As String, RFC As String

    AsegurarEncabezados
    razon = Trim(UserForm1.txtRazonSocial.Value)
    RFC = Trim(UserForm1.txtRFC.Value)
    If razon = "" And RFC = "" Then Exit Function

    Fila = BuscarEmpresa(RFC, razon)
    If Fila = 0 Then
        If MsgBox("La empresa """ & razon & """ (RFC: " & IIf(RFC = "", "sin RFC", RFC) & _
                  ") no está registrada." & vbCrLf & "¿Deseas registrarla ahora?", _
                  vbQuestion + vbYesNo, "CALEIDO") = vbNo Then Exit Function
        Fila = FilaEmpresaPorID(frmEmpresa.Registrar(razon, RFC, _
                   UserForm1.txtEmail.Value, UserForm1.txtDomFiscal.Value))
        If Fila = 0 Then Exit Function

        'El pedido se guarda con los datos tal como quedaron registrados
        UserForm1.txtRazonSocial.Value = Sheets("Empresas").Cells(Fila, "B").Value
        UserForm1.txtRFC.Value = Sheets("Empresas").Cells(Fila, "C").Value
        UserForm1.txtDomFiscal.Value = DomicilioEmpresa(Fila)
    End If

    ResolverEmpresa = EmpresaIDDeFila(Fila)
End Function
