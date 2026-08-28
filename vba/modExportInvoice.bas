Attribute VB_Name = "modExportInvoice"
Option Explicit

Sub transferInvoiceToTemplate()

    ''set sheet variable
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
        
    Set wsInput = Sheet7
    Set wsOutput = Sheet1
    
    ''find header column in input ws
    Dim skuColIn As Long
    Dim descColIn As Long
    Dim unitColIn As Long
    Dim qtyColIn As Long
    Dim priceColIn As Long
    Dim amColIn As Long
    
    skuColIn = FindHeaderColumn(wsInput, 32, "SKU")
    descColIn = FindHeaderColumn(wsInput, 32, "Description")
    unitColIn = FindHeaderColumn(wsInput, 32, "Unit")
    qtyColIn = FindHeaderColumn(wsInput, 32, "Qty")
    priceColIn = FindHeaderColumn(wsInput, 32, "Price/Unit")
    amColIn = FindHeaderColumn(wsInput, 32, "Amount")
    
    ''find header column in output ws
    
    Dim skuColOut As Long
    Dim descColOut As Long
    Dim qtyColOut As Long
    Dim priceUnitColOut As Long
    Dim amColOut As Long
    
    skuColOut = FindHeaderColumn(wsOutput, 20, "SKU")
    descColOut = FindHeaderColumn(wsOutput, 20, "Description")
    qtyColOut = FindHeaderColumn(wsOutput, 20, "Qty")
    priceUnitColOut = FindHeaderColumn(wsOutput, 20, "Price/Unit")
    amColOut = FindHeaderColumn(wsInput, 20, "Amount")
    
    ''clear template Invoice
    With Sheet1
    
        Range("I7:J10").ClearContents
        Range("E12").ClearContents
        Range("C21:H35").ClearContents
    
    End With
    
    ''transfer invoice data from wsInput to wsOutput
    
    With wsOutput
    
        .Range("I7").Value = wsInput.Range("V8").Value
        .Range("I8").Value = wsInput.Range("V9").Value
        .Range("I9").Value = wsInput.Range("V10").Value
        .Range("I10").Value = wsInput.Range("V11").Value
        .Range("E12").Value = wsInput.Range("K16").Value
    
    End With
        


End Sub
