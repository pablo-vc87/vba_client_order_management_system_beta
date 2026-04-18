VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiaCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Dia As MSForms.Label
Attribute Dia.VB_VarHelpID = -1
Private ParentForm As Object

Public Sub Inicializar(lbl As MSForms.Label, frm As Object)
    Set Dia = lbl
    Set ParentForm = frm
End Sub

Private Sub Dia_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Dia.Caption <> "-" Then
        Dim d As Long, m As Long, a As Long
        
        d = CLng(Dia.Caption)
        m = CLng(ParentForm.cboMes.Value)
        a = CLng(ParentForm.lblAno.Caption)
        
        Unload ParentForm
        ModuloCalendario.RecibeLaFecha d, m, a
    End If
End Sub

Private Sub Dia_MouseMove(ByVal Button As Integer, _
                          ByVal Shift As Integer, _
                          ByVal X As Single, _
                          ByVal Y As Single)
    ModuloCalendario.MarcarDia Dia
End Sub

