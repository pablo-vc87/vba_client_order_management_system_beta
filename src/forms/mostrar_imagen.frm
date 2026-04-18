VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mostrar_imagen 
   Caption         =   "CALEIDO | Render"
   ClientHeight    =   29748
   ClientLeft      =   336
   ClientTop       =   1440
   ClientWidth     =   54768
   OleObjectBlob   =   "mostrar_imagen.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "mostrar_imagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Activate()
txtrutaimg.Locked = True
Me.Height = 601.2
Me.Width = 1133.4
Me.StartUpPosition = 2

On Error GoTo malemerga
imggrande.Picture = LoadPicture(ActiveCell)
imggrande.PictureSizeMode = fmPictureSizeModeZoom

Exit Sub
malemerga:
MsgBox "No se encontró la imagen", , "CALEIDO"
End Sub

