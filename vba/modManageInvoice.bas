Attribute VB_Name = "modManageInvoice"
Option Explicit

Public Sub updateManageInvoiceScrollBarMax()

    Dim ws As Worksheet
    Set ws = Sheet11

    Dim countDataRow As Long

    countDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 9
    MsgBox (countDataRow)
    If (countDataRow < 13) Then
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = 1
            
        Else
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = countDataRow - 13
            
    End If
    
    
End Sub


Public Sub filterDisplay()

    Dim ws As Worksheet
    Set ws = Sheet11
    
    If ws.Range("W10") = "" Then
        ws.Range("W10").Value = "All"
    End If
    

End Sub

Sub applyFilter()

    With Sheet11
    
        .Range("H10").Value = .Range("Z10").Value
        .Range("K10").Value = .Range("AF10").Value
        
        If .Range("AC10").Value = "All" Then
        
            .Range("I10").Value = ""
        Else
            .Range("I10").Value = .Range("AC10").Value
        End If
        
    
    End With
    
    
End Sub

Sub clearFilter()

    With Sheet11
        .Range("Z10").Value = "All"
        .Range("AC10").Value = "All"
        .Range("AF10").Value = "All"
    
    Call applyFilter
    
    End With
    
    
End Sub


