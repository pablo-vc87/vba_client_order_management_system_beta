Attribute VB_Name = "ModuloFacturacion"
Option Explicit

' Prefactura y complemento de pago (REP) de un pedido, en Word y PDF.
' NO son comprobantes fiscales: reúnen los datos que se capturan en el portal del SAT o
' del PAC para timbrar el CFDI.
'   - Pago en una sola exhibición: prefactura con método PUE.
'   - Anticipo o parcialidades: prefactura con método PPD (forma 99) y un complemento de
'     pago por cada abono (hoja Pagos). Requiere el UUID de la factura ya timbrada.
' Hojas: Config (datos del emisor) y Pagos. Las claves de producto/servicio y de unidad
' del SAT se capturan a mano al timbrar.
' Columnas de Pedidos: P EmpresaID, Q FormaPago, R MetodoPago, S UsoCFDI, T UUIDFactura.

Private Const PLANTILLA_PF As String = "Plantilla_Prefactura.docx"
Private Const PLANTILLA_REP As String = "Plantilla_Complemento_Pago.docx"
Private Const CARPETA_FACTURAS As String = "facturas"
Private Const TASA_IVA As Double = 0.16

' ===============================================
' Botones / macros
' ===============================================
Public Sub Prefactura_Pedido()
    Dim r As String
    r = InputBox("Número de pedido a facturar:", "CALEIDO")
    If IsNumeric(r) Then GenerarPrefactura CLng(r)
End Sub

Public Sub Pago_Pedido()
    Dim r As String
    r = InputBox("Número de pedido que se está pagando:", "CALEIDO")
    If IsNumeric(r) Then RegistrarPago CLng(r)
End Sub

' ===============================================
' UNA SOLA VEZ: crea las hojas y columnas de facturación (no borra nada que ya exista)
' ===============================================
Public Sub PrepararFacturacion()
    Dim ws As Worksheet

    Set ws = ObtenerHoja("Config")
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:B1").Value = Array("Campo", "Valor")
        ws.Range("A2").Value = "RFC"
        ws.Range("A3").Value = "RazonSocial"
        ws.Range("A4").Value = "RegimenFiscal"
        ws.Range("A5").Value = "CPExpedicion"
        ws.Range("B2:B5").NumberFormat = "@"
        ws.Range("C2").Value = "<- Datos del emisor (CALEIDO). Régimen: solo la clave, ej. 612. C.P.: el del domicilio fiscal."
    End If

    Set ws = ObtenerHoja("Pagos")
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:H1").Value = Array("PagoID", "PedidoID", "FechaPago", "Monto", "FormaPago", _
                                        "Parcialidad", "SaldoAnterior", "SaldoInsoluto")
        ws.Columns("C").NumberFormat = "dd/mm/yyyy"
        ws.Columns("E").NumberFormat = "@"
    End If

    With Sheets("Pedidos")
        If .Range("P1").Value = "" Then .Range("P1").Value = "EmpresaID"
        .Range("Q1:T1").Value = Array("FormaPago", "MetodoPago", "UsoCFDI", "UUIDFactura")
        .Columns("Q").NumberFormat = "@"
    End With

    ThisWorkbook.Save
    MsgBox "Listo. Llena los datos del emisor en la hoja Config.", _
           vbInformation, "CALEIDO"
End Sub

' ===============================================
' PREFACTURA
' ===============================================
Public Sub GenerarPrefactura(ByVal PedidoID As Long)
    Dim filaPed As Long, filaEmp As Long, ex As Object
    Dim forma As String, metodo As String, uso As String

    If Not ListoParaFacturar(PedidoID, filaPed, filaEmp) Then Exit Sub

    With Sheets("Pedidos")
        forma = Trim(.Cells(filaPed, "Q").Value)
        metodo = Trim(.Cells(filaPed, "R").Value)
        uso = Trim(.Cells(filaPed, "S").Value)
        If uso = "" Then uso = Trim(Sheets("Empresas").Cells(filaEmp, "M").Value)

        If Not frmFacturar.PedirDatosFactura(forma, metodo, uso, Numero(.Cells(filaPed, "O").Value)) Then Exit Sub

        .Cells(filaPed, "Q").NumberFormat = "@"
        .Cells(filaPed, "Q").Value = forma
        .Cells(filaPed, "R").Value = metodo
        .Cells(filaPed, "S").Value = uso
    End With

    Set ex = DatosEmisorReceptor(filaEmp, uso)
    ex("FECHA_FACTURA") = Format(Date, "dd/mm/yyyy")
    ex("FORMA_PAGO") = TextoCatalogo(ListaFormasPago(True), forma)
    ex("METODO_PAGO") = TextoCatalogo(ListaMetodosPago(), metodo)
    If metodo = "PPD" Then
        ex("NOTA_METODO") = "Pago en parcialidades o diferido: cada pago se acredita con un complemento de pago (REP)."
    Else
        ex("NOTA_METODO") = ""
    End If

    GenerarDocumento PedidoID, PLANTILLA_PF, "PF", ex, CARPETA_FACTURAS
End Sub

' ===============================================
' PAGO (anticipo o parcialidad) + COMPLEMENTO DE PAGO
' ===============================================
Public Sub RegistrarPago(ByVal PedidoID As Long)
    Dim filaPed As Long, filaEmp As Long, ws As Worksheet, Fila As Long, ex As Object
    Dim total As Double, pagado As Double, saldo As Double, n As Long
    Dim fecha As Date, monto As Double, forma As String, uuid As String

    If Not ListoParaFacturar(PedidoID, filaPed, filaEmp) Then Exit Sub

    If Trim(Sheets("Pedidos").Cells(filaPed, "R").Value) <> "PPD" Then
        MsgBox "El complemento de pago solo aplica a facturas con método PPD." & vbCrLf & _
               "Genera primero la prefactura del pedido con método PPD.", vbExclamation, "CALEIDO"
        Exit Sub
    End If

    total = Round(Numero(Sheets("Pedidos").Cells(filaPed, "O").Value), 2)
    PagosDelPedido PedidoID, pagado, n
    saldo = Round(total - pagado, 2)
    If saldo <= 0 Then
        MsgBox "El pedido " & PedidoID & " ya está pagado por completo.", vbInformation, "CALEIDO"
        Exit Sub
    End If

    fecha = Date
    monto = saldo
    forma = "03"
    uuid = Trim(Sheets("Pedidos").Cells(filaPed, "T").Value)
    If Not frmFacturar.PedirPago(total, pagado, saldo, fecha, monto, forma, uuid) Then Exit Sub

    Sheets("Pedidos").Cells(filaPed, "T").Value = uuid

    'Registro del pago
    Set ws = Sheets("Pagos")
    Fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(Fila, "A").Value = Application.WorksheetFunction.Max(ws.Columns("A")) + 1
    ws.Cells(Fila, "B").Value = PedidoID
    ws.Cells(Fila, "C").Value = fecha
    ws.Cells(Fila, "D").Value = monto
    ws.Cells(Fila, "E").NumberFormat = "@"
    ws.Cells(Fila, "E").Value = forma
    ws.Cells(Fila, "F").Value = n + 1
    ws.Cells(Fila, "G").Value = saldo
    ws.Cells(Fila, "H").Value = Round(saldo - monto, 2)

    'Complemento de pago
    Set ex = DatosEmisorReceptor(filaEmp, "CP01")
    ex("FECHA_EMISION") = Format(Date, "dd/mm/yyyy")
    ex("FECHA_PAGO") = Format(fecha, "dd/mm/yyyy")
    ex("FORMA_PAGO") = TextoCatalogo(ListaFormasPago(False), forma)
    ex("MONTO_PAGO") = Format(monto, "#,##0.00")
    ex("MONTO_LETRA") = ImporteConLetra(monto)
    ex("PARCIALIDAD") = CStr(n + 1)
    ex("SALDO_ANTERIOR") = Format(saldo, "#,##0.00")
    ex("SALDO_INSOLUTO") = Format(saldo - monto, "#,##0.00")
    ex("BASE_PAGO") = Format(monto / (1 + TASA_IVA), "#,##0.00")
    ex("IVA_PAGO") = Format(monto - monto / (1 + TASA_IVA), "#,##0.00")
    If uuid <> "" Then
        ex("UUID_FACTURA") = UCase(uuid)
    Else
        ex("UUID_FACTURA") = "(pendiente: UUID de la factura timbrada)"
    End If

    ThisWorkbook.Save
    GenerarDocumento PedidoID, PLANTILLA_REP, "REP-P" & (n + 1), ex, CARPETA_FACTURAS
End Sub

' Revisa que existan el pedido, su empresa y los datos del emisor
Private Function ListoParaFacturar(ByVal PedidoID As Long, ByRef filaPed As Long, _
                                   ByRef filaEmp As Long) As Boolean
    If Not HojaExiste("Config") Or Not HojaExiste("Pagos") Then
        MsgBox "Primero corre la macro PrepararFacturacion (una sola vez).", vbExclamation, "CALEIDO"
        Exit Function
    End If
    Dim campo As Variant, faltan As String
    For Each campo In Array("RFC", "RazonSocial", "RegimenFiscal", "CPExpedicion")
        If DatoConfig(campo) = "" Then faltan = faltan & "- " & campo & vbCrLf
    Next campo
    If faltan <> "" Then
        MsgBox "Faltan datos del emisor en la hoja Config (columna B, junto a cada campo):" & _
               vbCrLf & faltan, vbExclamation, "CALEIDO"
        Exit Function
    End If
    If TextoCatalogo(ListaRegimenes(), DatoConfig("RegimenFiscal")) = DatoConfig("RegimenFiscal") Then
        MsgBox "Config!B4 (RegimenFiscal) debe ser la clave del régimen de CALEIDO, ej. 612 o 626." & _
               vbCrLf & "Hoy dice: " & DatoConfig("RegimenFiscal"), vbExclamation, "CALEIDO"
        Exit Function
    End If
    If Not DatoConfig("CPExpedicion") Like "#####" Then
        MsgBox "Config!B5 (CPExpedicion) debe ser el código postal de 5 dígitos del domicilio fiscal." & _
               vbCrLf & "Hoy dice: " & DatoConfig("CPExpedicion"), vbExclamation, "CALEIDO"
        Exit Function
    End If

    filaPed = FilaPedido(PedidoID)
    If filaPed = 0 Then
        MsgBox "No existe el pedido " & PedidoID, vbExclamation, "CALEIDO"
        Exit Function
    End If

    filaEmp = FilaEmpresaPorID(Val(Sheets("Pedidos").Cells(filaPed, "P").Value))
    If filaEmp = 0 Then
        MsgBox "El pedido " & PedidoID & " no está ligado a una empresa." & vbCrLf & _
               "Ábrelo en Buscar > Modificar, elige o registra la empresa y guarda.", _
               vbExclamation, "CALEIDO"
        Exit Function
    End If

    ListoParaFacturar = True
End Function

' Marcadores del emisor (hoja Config) y del receptor (hoja Empresas)
Private Function DatosEmisorReceptor(ByVal filaEmp As Long, ByVal uso As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d("EMISOR_RAZON") = DatoConfig("RazonSocial")
    d("EMISOR_RFC") = DatoConfig("RFC")
    d("EMISOR_REGIMEN") = TextoCatalogo(ListaRegimenes(), DatoConfig("RegimenFiscal"))
    d("EMISOR_CP") = DatoConfig("CPExpedicion")

    With Sheets("Empresas")
        d("RECEPTOR_RAZON") = Trim(.Cells(filaEmp, "B").Value)
        d("RECEPTOR_RFC") = Trim(.Cells(filaEmp, "C").Value)
        d("RECEPTOR_CP") = Trim(.Cells(filaEmp, "H").Value)
        d("RECEPTOR_REGIMEN") = TextoCatalogo(ListaRegimenes(), Trim(.Cells(filaEmp, "L").Value))
        d("RECEPTOR_DOMICILIO") = DomicilioEmpresa(filaEmp)
        d("EMAIL_FACTURACION") = Trim(.Cells(filaEmp, "N").Value)
    End With
    d("USO_CFDI") = TextoCatalogo(ListaUsosCFDI(), uso)

    Set DatosEmisorReceptor = d
End Function

' Suma y número de pagos ya registrados de un pedido
Private Sub PagosDelPedido(ByVal PedidoID As Long, ByRef pagado As Double, ByRef n As Long)
    Dim ws As Worksheet, i As Long
    Set ws = Sheets("Pagos")
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If Val(ws.Cells(i, "B").Value) = PedidoID Then
            pagado = pagado + Numero(ws.Cells(i, "D").Value)
            n = n + 1
        End If
    Next i
End Sub

' ===============================================
' Datos de la hoja Config y descripción de conceptos
' ===============================================
Public Function DatoConfig(ByVal campo As String) As String
    Dim c As Range
    Set c = Sheets("Config").Columns("A").Find(campo, LookAt:=xlWhole)
    If Not c Is Nothing Then DatoConfig = Trim(CStr(c.Offset(0, 1).Value))
End Function

' Descripción del concepto en la prefactura: material y tamaño (ej. "Playera algodón - 10 x 10 cm")
Public Function DescripcionConcepto(ByVal material As String, ByVal tamano As String) As String
    If material <> "" And tamano <> "" Then
        DescripcionConcepto = material & " - " & tamano
    Else
        DescripcionConcepto = material & tamano
    End If
End Function

' Importe de una celda o campo (CDbl y no Val: Val trunca si Windows usa coma decimal)
Public Function Numero(ByVal v As Variant) As Double
    If IsNumeric(v) Then Numero = CDbl(v)
End Function

Private Function HojaExiste(ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function

' ===============================================
' Catálogos del SAT ("clave - descripción")
' ===============================================
Public Function ListaFormasPago(ByVal incluirPorDefinir As Boolean) As Variant
    If incluirPorDefinir Then
        ListaFormasPago = Array("01 - Efectivo", "02 - Cheque nominativo", _
            "03 - Transferencia electrónica de fondos", "04 - Tarjeta de crédito", _
            "28 - Tarjeta de débito", "99 - Por definir")
    Else
        ListaFormasPago = Array("01 - Efectivo", "02 - Cheque nominativo", _
            "03 - Transferencia electrónica de fondos", "04 - Tarjeta de crédito", _
            "28 - Tarjeta de débito")
    End If
End Function

Public Function ListaMetodosPago() As Variant
    ListaMetodosPago = Array("PUE - Pago en una sola exhibición", _
                             "PPD - Pago en parcialidades o diferido")
End Function

' Usos para facturas de ingreso (CP01 - Pagos solo va en el complemento de pago)
Public Function ListaUsosCFDI() As Variant
    ListaUsosCFDI = Array("G01 - Adquisición de mercancías", _
        "G02 - Devoluciones, descuentos o bonificaciones", "G03 - Gastos en general", _
        "S01 - Sin efectos fiscales")
End Function

Public Function ListaRegimenes() As Variant
    ListaRegimenes = Array("601 - General de Ley Personas Morales", _
        "603 - Personas Morales con Fines no Lucrativos", "605 - Sueldos y Salarios", _
        "606 - Arrendamiento", "610 - Residentes en el Extranjero", _
        "612 - Personas Físicas con Actividades Empresariales y Profesionales", _
        "616 - Sin obligaciones fiscales", "621 - Incorporación Fiscal", _
        "623 - Opcional para Grupos de Sociedades", "624 - Coordinados", _
        "625 - Actividades Empresariales con ingresos a través de Plataformas Tecnológicas", _
        "626 - Régimen Simplificado de Confianza")
End Function

' Clave de un elemento "clave - descripción"
Public Function ClaveCatalogo(ByVal elemento As String) As String
    ClaveCatalogo = Trim(Split(elemento & " - ", " - ")(0))
End Function

' "clave - descripción" a partir de la clave (si no está en la lista, la clave sola)
Public Function TextoCatalogo(ByVal lista As Variant, ByVal clave As String) As String
    Dim e As Variant
    clave = Trim(clave)
    TextoCatalogo = clave
    For Each e In lista
        If ClaveCatalogo(e) = clave Then
            TextoCatalogo = e
            Exit Function
        End If
    Next e
End Function

' ===============================================
' Importe con letra: 1500.5 -> "MIL QUINIENTOS PESOS 50/100 M.N." (hasta 999,999,999.99)
' ===============================================
Public Function ImporteConLetra(ByVal monto As Double) As String
    Dim enteros As Double, centavos As Long

    monto = Round(monto, 2)
    enteros = Int(monto)
    centavos = CLng(Round((monto - enteros) * 100, 0))
    If centavos = 100 Then
        enteros = enteros + 1
        centavos = 0
    End If

    'Millones cerrados: "un millón de pesos"
    If enteros >= 1000000# And enteros - Int(enteros / 1000000#) * 1000000# = 0 Then
        ImporteConLetra = UCase(NumeroEnLetras(enteros)) & " DE PESOS "
    Else
        ImporteConLetra = UCase(NumeroEnLetras(enteros)) & IIf(enteros = 1, " PESO ", " PESOS ")
    End If
    ImporteConLetra = ImporteConLetra & Format(centavos, "00") & "/100 M.N."
End Function

Private Function NumeroEnLetras(ByVal n As Double) As String
    Dim millones As Long, miles As Long, resto As Long, s As String

    If n = 0 Then
        NumeroEnLetras = "cero"
        Exit Function
    End If
    millones = Int(n / 1000000#)
    miles = Int((n - millones * 1000000#) / 1000#)
    resto = n - millones * 1000000# - miles * 1000#

    If millones = 1 Then
        s = "un millón"
    ElseIf millones > 1 Then
        s = Centenas(millones, True) & " millones"
    End If
    If miles = 1 Then
        s = s & " mil"
    ElseIf miles > 1 Then
        s = s & " " & Centenas(miles, True) & " mil"
    End If
    If resto > 0 Then s = s & " " & Centenas(resto, True)       'antes de "pesos": un, veintiún

    NumeroEnLetras = Trim(s)
End Function

' 1 a 999. apocope = True antes de un sustantivo: mil, millones, pesos (uno -> un, veintiuno -> veintiún)
Private Function Centenas(ByVal n As Long, ByVal apocope As Boolean) As String
    Dim c As Long, d As Long, u As Long, s As String
    Dim unidades As Variant, especiales As Variant, veintes As Variant
    Dim decenas As Variant, cientos As Variant

    unidades = Array("", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve")
    especiales = Array("diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", _
                       "diecisiete", "dieciocho", "diecinueve")
    veintes = Array("veinte", "veintiuno", "veintidós", "veintitrés", "veinticuatro", _
                    "veinticinco", "veintiséis", "veintisiete", "veintiocho", "veintinueve")
    decenas = Array("", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", _
                    "ochenta", "noventa")
    cientos = Array("", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", _
                    "seiscientos", "setecientos", "ochocientos", "novecientos")

    If n = 100 Then
        Centenas = "cien"
        Exit Function
    End If

    c = n \ 100
    d = (n Mod 100) \ 10
    u = n Mod 10

    s = cientos(c)
    Select Case d
        Case 0
            If u > 0 Then s = s & " " & unidades(u)
        Case 1
            s = s & " " & especiales(u)
        Case 2
            s = s & " " & veintes(u)
        Case Else
            s = s & " " & decenas(d)
            If u > 0 Then s = s & " y " & unidades(u)
    End Select
    s = Trim(s)

    If apocope Then
        If Right(s, 9) = "veintiuno" Then
            s = Left(s, Len(s) - 9) & "veintiún"
        ElseIf Right(s, 3) = "uno" Then
            s = Left(s, Len(s) - 1)
        End If
    End If

    Centenas = s
End Function
