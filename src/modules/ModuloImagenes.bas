Attribute VB_Name = "ModuloImagenes"
Option Explicit

' Las imágenes de los productos se copian a una subcarpeta junto al libro y en
' Detalle_Pedidos se guarda la ruta RELATIVA (ej. "imagenes\logo.jpg"). Así el libro
' y su carpeta se pueden mover a otra computadora sin que se rompan las rutas.
Public Const CARPETA_IMAGENES As String = "imagenes"

' ===============================================
' Carpeta absoluta de imágenes (la crea si no existe)
' ===============================================
Public Function CarpetaImagenes() As String
    CarpetaImagenes = ThisWorkbook.Path & "\" & CARPETA_IMAGENES
    If Dir(CarpetaImagenes, vbDirectory) = "" Then MkDir CarpetaImagenes
End Function

' ===============================================
' Copia la imagen elegida a la carpeta del libro y devuelve su ruta relativa.
' Si ya está en la carpeta, no la copia otra vez. Si existe otra con el mismo
' nombre, agrega _1, _2...
' ===============================================
Public Function CopiarImagenAlLibro(ByVal rutaOrigen As String) As String
    Dim carpeta As String, nombre As String, base As String, ext As String
    Dim destino As String, n As Long, p As Long

    carpeta = CarpetaImagenes()
    nombre = Mid(rutaOrigen, InStrRev(rutaOrigen, "\") + 1)

    If LCase(rutaOrigen) = LCase(carpeta & "\" & nombre) Then
        CopiarImagenAlLibro = CARPETA_IMAGENES & "\" & nombre
        Exit Function
    End If

    p = InStrRev(nombre, ".")
    If p > 0 Then
        base = Left(nombre, p - 1)
        ext = Mid(nombre, p)
    Else
        base = nombre
    End If

    destino = carpeta & "\" & nombre
    Do While Dir(destino) <> ""
        'Mismo nombre y mismo tamaño: es la misma imagen, ya copiada antes
        If FileLen(destino) = FileLen(rutaOrigen) Then
            CopiarImagenAlLibro = CARPETA_IMAGENES & "\" & nombre
            Exit Function
        End If
        n = n + 1
        nombre = base & "_" & n & ext
        destino = carpeta & "\" & nombre
    Loop

    FileCopy rutaOrigen, destino
    CopiarImagenAlLibro = CARPETA_IMAGENES & "\" & nombre
End Function

' ===============================================
' Convierte la ruta guardada en una ruta que se pueda abrir:
' - relativa ("imagenes\x.jpg")  -> carpeta del libro + ruta
' - absoluta vieja (C:\...\x.jpg) -> se usa si existe; si no, se busca
'   el mismo nombre en la carpeta de imágenes (pedidos de otra computadora)
' ===============================================
Public Function RutaAbsolutaImagen(ByVal ruta As String) As String
    Dim alterna As String

    ruta = Trim(ruta)
    If ruta = "" Then Exit Function

    If Mid(ruta, 2, 1) = ":" Or Left(ruta, 2) = "\\" Then
        RutaAbsolutaImagen = ruta
        If Dir(ruta) = "" Then
            alterna = ThisWorkbook.Path & "\" & CARPETA_IMAGENES & "\" & _
                      Mid(ruta, InStrRev(ruta, "\") + 1)
            If Dir(alterna) <> "" Then RutaAbsolutaImagen = alterna
        End If
    Else
        RutaAbsolutaImagen = ThisWorkbook.Path & "\" & ruta
    End If
End Function

' ===============================================
' Abre una imagen para un control Image del formulario.
' LoadPicture solo abre BMP, GIF, ICO, WMF y JPG no progresivo; WIA (Windows Image
' Acquisition) abre también PNG y JPG progresivo. LoadPicture queda de respaldo.
' Devuelve Nothing si no se pudo abrir.
' ===============================================
Public Function CargarPicture(ByVal ruta As String) As Object
    Dim wia As Object

    If ruta = "" Then Exit Function
    If Dir(ruta) = "" Then Exit Function

    On Error Resume Next
    Set wia = CreateObject("WIA.ImageFile")
    wia.LoadFile ruta
    Set CargarPicture = wia.FileData.Picture
    If CargarPicture Is Nothing Then
        Err.Clear
        Set CargarPicture = LoadPicture(ruta)
    End If
    On Error GoTo 0
End Function

' ===============================================
' UNA SOLA VEZ: copia las imágenes de los pedidos existentes (rutas absolutas)
' a la carpeta del libro y cambia Detalle_Pedidos!K a rutas relativas.
' Hazlo en esta computadora, donde las rutas viejas todavía existen.
' ===============================================
Public Sub MigrarImagenesExistentes()
    Dim ws As Worksheet, i As Long, ultFila As Long
    Dim ruta As String, copiadas As Long, faltantes As Long

    Set ws = ThisWorkbook.Sheets("Detalle_Pedidos")
    ultFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultFila
        ruta = Trim(ws.Cells(i, "K").Value)
        If Mid(ruta, 2, 1) = ":" Or Left(ruta, 2) = "\\" Then
            If Dir(ruta) <> "" Then
                ws.Cells(i, "K").Value = CopiarImagenAlLibro(ruta)
                copiadas = copiadas + 1
            Else
                faltantes = faltantes + 1
            End If
        End If
    Next i

    ThisWorkbook.Save
    MsgBox "Imágenes copiadas: " & copiadas & vbCrLf & _
           "No encontradas (se dejaron igual): " & faltantes, vbInformation, "CALEIDO"
End Sub
