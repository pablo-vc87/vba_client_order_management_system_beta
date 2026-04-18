VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1
Public ParentForm As Object
Public Index As Long

'========================
' CLICK BOTONES
'========================
Private Sub Btn_Click()

    ' Seguridad: solo botones válidos
    If Btn Is Nothing Then Exit Sub
    If Not TypeOf Btn Is MSForms.CommandButton Then Exit Sub

    ' ========================
    ' BOTÓN CALENDARIO
    ' ========================
    If Btn.Name Like "cmdFechaRec*" Then

        ' Validar que el textbox exista
        On Error Resume Next
        Dim tb As MSForms.TextBox
        Set tb = ParentForm.frProductos.Controls("txtFechaRec" & Index)
        On Error GoTo 0

        If tb Is Nothing Then Exit Sub

        Set ModuloCalendario.TextBoxDestino = tb
        frmCalendario.Show

        Exit Sub
    End If

    ' ========================
    ' BOTÓN IMAGEN
    ' ========================
    If Btn.Name Like "cmdImg*" Then

        Dim fd As FileDialog
        'Dim ruta As String ' nvo
        Set fd = Application.FileDialog(msoFileDialogFilePicker)

        With fd
            .Title = "Seleccionar imagen"
            .Filters.Clear
            .Filters.Add "Imágenes", "*.jpg;*.png;*.bmp;*.jpeg"
            .AllowMultiSelect = False
            
            If .Show Then
                ParentForm.frProductos.Controls("txtImg" & Index).Value = .SelectedItems(1)
                ParentForm.frProductos.Controls("imgPrev" & Index).Picture = _
                    LoadPicture(.SelectedItems(1))
                ' ParentForm.MostrarPreviewImagen Index
            End If
            'If .Show <> -1 Then Exit Sub
            'ruta = .SelectedItems(1)
        End With
        ' Guardar ruta NVO
        'UserForm1.frProductos.Controls("txtImg" & Index).Value = ruta
        ' MOSTRAR PREVIEW NVO
        'UserForm1.MostrarPreviewImagen Index
    End If

End Sub

'========================
' VALIDAR TEXTO NUMÉRICO
'========================
Private Sub Txt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    ' Solo números, punto y backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57 _
        Or KeyAscii = 46 _
        Or KeyAscii = 8) Then
        KeyAscii = 0
    End If

    ' Solo un punto decimal
    If KeyAscii = 46 Then
        If InStr(Txt.Text, ".") > 0 Then KeyAscii = 0
    End If

End Sub

