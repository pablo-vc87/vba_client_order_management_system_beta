Attribute VB_Name = "ModuloCalendario"
Option Explicit
Option Private Module
Private SenalCambioMes As Long
Public TextBoxDestino As MSForms.TextBox


Public Sub RecibeLaFecha(Dia As Long, Mes As Long, Ano As Long)
    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial((VBA.CInt(Ano)), (VBA.CInt(Mes)), (VBA.CInt(Dia)))
    
'INDICAR UBICACIÓN DE LA FECHA
    If Not TextBoxDestino Is Nothing Then
        TextBoxDestino.Value = FechaRecibida
    End If
    
    
    
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************

Public Sub InicializaFormularioCalendario()
    SenalCambioMes = 1

    frmCalendario.cboMes.Clear

    With frmCalendario.cboMes
        .AddItem 1: .List(0, 1) = "enero"
        .AddItem 2: .List(1, 1) = "febrero"
        .AddItem 3: .List(2, 1) = "marzo"
        .AddItem 4: .List(3, 1) = "abril"
        .AddItem 5: .List(4, 1) = "mayo"
        .AddItem 6: .List(5, 1) = "junio"
        .AddItem 7: .List(6, 1) = "julio"
        .AddItem 8: .List(7, 1) = "agosto"
        .AddItem 9: .List(8, 1) = "septiembre"
        .AddItem 10: .List(9, 1) = "octubre"
        .AddItem 11: .List(10, 1) = "noviembre"
        .AddItem 12: .List(11, 1) = "diciembre"
    End With

    frmCalendario.cboMes.ListIndex = Month(Date) - 1
    frmCalendario.spbAño.Value = Year(Date)
    frmCalendario.lblAno.Caption = Year(Date)

    Call CargarLosDias(Year(Date), Month(Date))

    frmCalendario.lblHoy.Caption = Format(Date, "yyyy-mm-dd")
End Sub

Public Sub CargarLosDias(Ano As Long, Mes As Long)
    Dim FechaDelPrimerDia As Date
    Dim FechaDelUltimoDia As Date
    Dim DiaSemanaPrimerDia As Long
    Dim VariableControl As Control
    Dim Contador As Long
    
    FechaDelPrimerDia = VBA.DateSerial(Ano, Mes, 1)
    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, Mes, 1), 0)
    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
    Contador = 1
    
    For Each VariableControl In frmCalendario.mrcDias.Controls
        VariableControl.Caption = "-"
        If VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia) Then
            VariableControl.Caption = Contador
            Contador = Contador + 1
        End If
    Next VariableControl
End Sub

Public Sub CambioDeMes()
    If SenalCambioMes > 1 Then
        Dim MesEnElCombo As Long, AnoEnElLabel As Long
        
        If Not (IsNull(frmCalendario.cboMes.Value)) And Not (IsNull(frmCalendario.lblAno.Caption)) Then
            MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
            AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
            Call ModuloCalendario.DesmarcarDias
            Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
        End If
    End If
    SenalCambioMes = SenalCambioMes + 1
End Sub

Public Sub CambioDeAno()
    Dim MesEnElCombo As Long, AnoEnElLabel As Long
    
    frmCalendario.lblAno.Caption = frmCalendario.spbAño.Value
    
    MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
    
End Sub

Public Sub UnClickEnHoyEs()
    Dim Mes As Long, Ano As Long
    Dim FechaActual As Date
    
    FechaActual = VBA.CDate(frmCalendario.lblHoy.Caption)
    Mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.Year(FechaActual))
    
    frmCalendario.lblAno.Caption = Ano
    frmCalendario.cboMes.ListIndex = Mes - 1
    frmCalendario.spbAño.Value = Ano
    frmCalendario.spbAño.SetFocus
    
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(Ano, Mes)
    
End Sub

Sub SalirConEscape()
    Unload frmCalendario
End Sub

Sub MarcarDia(ControlDeEtiqueta As Control)
    Call ModuloCalendario.DesmarcarDias
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub

Sub DesmarcarDias()
    Dim ControlEtiqueta As Control
    
    For Each ControlEtiqueta In frmCalendario.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************

' Nota del autor -----------------------------------------------------------------------------

' Creado por Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube

' Versión 1.0 - 20 de julio de 2015


