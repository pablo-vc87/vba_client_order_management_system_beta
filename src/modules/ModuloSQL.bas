Attribute VB_Name = "ModuloSQL"
Option Explicit

'Ejecuta una consulta sobre este mismo libro y devuelve un Recordset desconectado
Public Function EjecutarSQL(ByVal sql As String) As Object
    Dim cn As Object, rs As Object

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & ThisWorkbook.FullName & ";" & _
            "Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3          'adUseClient
    On Error GoTo Fallo
    rs.Open sql, cn, 3, 1          'adOpenStatic, adLockReadOnly
    On Error GoTo 0
    Set rs.ActiveConnection = Nothing
    cn.Close

    Set EjecutarSQL = rs
    Exit Function

Fallo:
    'Muestra el mensaje de ACE y el SQL que recibió; devuelve Nothing
    MsgBox "Error en la consulta:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & sql, _
           vbCritical, "SQL"
    Debug.Print sql
    cn.Close
End Function


'Devuelve la hoja; si no existe, la crea al final (evita el error 9)
Public Function ObtenerHoja(ByVal nombre As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    If ObtenerHoja Is Nothing Then
        Set ObtenerHoja = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ObtenerHoja.Name = nombre
    End If
End Function

'Escribe el resultado de una consulta (con encabezados) en una hoja
Public Sub VolcarResultado(ByVal sql As String, ByVal hoja As String)
    Dim rs As Object, ws As Worksheet, i As Long

    Set rs = EjecutarSQL(sql)
    If rs Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets(hoja)
    ws.Cells.Clear

    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    ws.Range("A2").CopyFromRecordset rs
    rs.Close
End Sub

Public Function TablaLista() As String
    TablaLista = "[Lista$]"
End Function


'Banco de pruebas: cambia la consulta, guarda el libro y ejecuta (F5)
Public Sub Probar()
    ThisWorkbook.Save
    VolcarResultado "SELECT * FROM [Pedidos$]", "Consulta"
End Sub

'Banco de pruebas con archivo: escribe la consulta en consulta.sql (misma carpeta
'que el libro), con saltos de línea y comentarios --, guarda y ejecuta (F5)
Public Sub ProbarArchivo()
    ThisWorkbook.Save
    VolcarResultado LeerSQL("consulta.sql"), "Consulta"
End Sub

'Lee un archivo .sql (UTF-8) de la carpeta del libro y quita las líneas de comentario,
'porque ACE no acepta comentarios --
Public Function LeerSQL(ByVal archivo As String) As String
    Dim st As Object, linea As Variant, sql As String

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                    'adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile ThisWorkbook.Path & "\" & archivo

    For Each linea In Split(Replace(st.ReadText, vbCr, ""), vbLf)
        If Left(Trim(linea), 2) <> "--" Then sql = sql & linea & vbLf
    Next linea
    st.Close

    LeerSQL = sql
End Function



'Arma "IIF(campo IS NULL, '', 'prefijo' & campo)"
Private Function Parte(ByVal Campo As String, ByVal prefijo As String) As String
    Parte = "IIF(" & Campo & " IS NULL, '', '" & prefijo & "' & " & Campo & ")"
End Function

'Domicilio fiscal en una línea a partir de las columnas de Empresas (a = alias de la tabla)
Public Function ExprDomicilio(ByVal a As String) As String
    ExprDomicilio = "Mid(" & Parte(a & ".Calle", ", ") & _
        " & " & Parte(a & ".NumExterior", " ") & _
        " & " & Parte(a & ".NumInterior", " Int. ") & _
        " & " & Parte(a & ".Colonia", ", Col. ") & _
        " & " & Parte(a & ".CP", ", C.P. ") & _
        " & " & Parte(a & ".Municipio", ", ") & _
        " & " & Parte(a & ".Estado", ", ") & ", 3)"
End Function

Public Function SQL_ContactosConEmpresa() As String
    SQL_ContactosConEmpresa = _
        "SELECT c.NombreCliente, c.Email, e.RazonSocial, e.RFC" & _
        ", " & ExprDomicilio("e") & " AS DomicilioFiscal" & _
        ", e.EmailFacturacion" & _
        " FROM [Clientes$] AS c" & _
        " INNER JOIN [Empresas$] AS e ON c.EmpresaID = e.EmpresaID" & _
        " ORDER BY e.RazonSocial, c.NombreCliente"
End Function

