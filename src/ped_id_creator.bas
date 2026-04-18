Attribute VB_Name = "ped_id_creator"
Public Function NuevoPedidoID() As Long
    With Sheets("Pedidos")
        If Application.WorksheetFunction.CountA(.Range("A:A")) = 0 Then
            NuevoPedidoID = 1
        Else
            NuevoPedidoID = Application.WorksheetFunction.Max(.Range("A:A")) + 1
        End If
    End With
End Function

