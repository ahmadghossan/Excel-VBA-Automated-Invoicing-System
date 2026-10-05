Attribute VB_Name = "modManageInvoice"
Option Explicit

Public Sub UpdateManageInvoiceScrollBarMax()

    Dim ws As Worksheet
    Set ws = Sheet11

    Dim countDataRow As Long
    Dim countAllData As Long
    
    countAllData = Sheet3.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1

    countDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 9
    If (countDataRow < 25) Then
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = 1
            
        Else
            ws.Shapes("Scroll Bar 5").ControlFormat.Max = countDataRow - 28
            
    End If
        
    
    ws.Range("Y12") = "Showing " & countDataRow - 4 & " out of " & countAllData & " Invoices"
            
    
End Sub


Public Sub FilterDisplay()

    Dim ws As Worksheet
    Set ws = Sheet11
    
    If ws.Range("W10") = "" Then
        ws.Range("W10").Value = "All"
    End If
    

End Sub

Public Sub ApplyFilter()

    With Sheet11
    
        .Range("H10").Value = .Range("Z10").Value
        .Range("K10").Value = .Range("AF10").Value
        
        If .Range("AC10").Value = "All" Then
        
            .Range("I10").Value = ""
        Else
            .Range("I10").Value = .Range("AC10").Value
        End If
       
    
    End With
    
    Call UpdateManageInvoiceScrollBarMax

End Sub

Public Sub ClearFilter()

    With Sheet11
        .Range("Z10").Value = "All"
        .Range("AC10").Value = "All"
        .Range("AF10").Value = "All"
    
    Call ApplyFilter
    
    End With
    
    Call UpdateManageInvoiceScrollBarMax
    
End Sub
Sub Export()

    Sheet7.Range("J6").Value = Sheet11.Range("J4").Value
    
  
    Call LoadInvoice
    Call exportToPdf

End Sub

Sub ViewEditInvoice()

    Sheet7.Range("J6").Value = Sheet11.Range("J4").Value
    Sheet7.Activate
    Call LoadInvoice

End Sub
