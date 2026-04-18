VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Seleccione una fecha"
   ClientHeight    =   71388.01
   ClientLeft      =   1296
   ClientTop       =   5352
   ClientWidth     =   1.31436e5
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Dias() As clsDiaCalendario

Private Sub UserForm_Initialize()

    With Me
        .Width = 220
        .Height = 240
        .StartUpPosition = 2 'Centro de pantalla
        .Font.Name = "Calibri"
        .Font.Size = 9
    End With

    InicializarDias
    ModuloCalendario.InicializaFormularioCalendario

End Sub

Private Sub InicializarDias()
    Dim ctrl As Control
    Dim i As Long
    
    ReDim Dias(1 To 42)
    i = 1
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            If ctrl.Name Like "lbl#" Or ctrl.Name Like "lbl##" Then
                Set Dias(i) = New clsDiaCalendario
                Dias(i).Inicializar ctrl, Me
                i = i + 1
                If i > 42 Then Exit For
            End If
        End If
    Next ctrl
End Sub

Private Sub cboMes_Click()
    ModuloCalendario.CambioDeMes
End Sub

Private Sub spbAño_Change()
    ModuloCalendario.CambioDeAno
End Sub

Private Sub lblHoy_Click()
    ModuloCalendario.UnClickEnHoyEs
End Sub

Private Sub cmdSalirConEscape_Click()
    ModuloCalendario.SalirConEscape
End Sub

