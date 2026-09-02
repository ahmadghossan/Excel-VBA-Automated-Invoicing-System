Attribute VB_Name = "modManageInvoice"
Option Explicit

Public Sub updateManageInvoiceScrollBarMax()

    Dim ws As Worksheet
    Set ws = Sheet11

    Dim countDataRow As Long

    countDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 16

    If (countDataRow < 13) Then
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = 1
            
        Else
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = countDataRow - 12
            
    End If

End Sub

