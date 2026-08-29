Attribute VB_Name = "modExportInvoice"
Option Explicit

Sub TransferInvoiceToTemplate()

    ''set sheet variable
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
        
    Set wsInput = Sheet7
    Set wsOutput = Sheet1
    
    ''clear template Invoice
    With Sheet1
    
        .Range("J7:J10").ClearContents
        .Range("E12").ClearContents
        .Range("C21:K35").ClearContents
    
    End With
    
    ''transfer invoice data from wsInput to wsOutput (from create invoice to template )
    
    With wsOutput
    
        .Range("J7").Value = wsInput.Range("V8").Value
        .Range("J8").Value = wsInput.Range("V9").Value
        .Range("J9").Value = wsInput.Range("V10").Value
        .Range("J10").Value = wsInput.Range("V11").Value
        .Range("E12").Value = wsInput.Range("K16").Value
    
    End With
        
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
    amColOut = FindHeaderColumn(wsOutput, 20, "Amount")

    

    ''transfer invoice transaction details from create invoice to template
    Dim i As Long
    Dim r As Long
    
    r = 21
    
    ''for every row from 33 to 47 (transaction detail row)
    For i = 33 To 47
    
        With wsInput
        ''if there is transaction detail for the row, then copy to template
        If .Cells(i, skuColIn) <> "" And .Cells(i, qtyColIn) <> "" Then
            
            wsOutput.Cells(r, skuColOut).Value = .Cells(i, skuColIn).Value
            wsOutput.Cells(r, descColOut).Value = .Cells(i, descColIn).Value
            wsOutput.Cells(r, qtyColOut).Value = .Cells(i, qtyColIn).Value
            wsOutput.Cells(r, priceUnitColOut).Value = .Cells(i, priceColIn).Value
            wsOutput.Cells(r, amColOut).Value = .Cells(i, amColIn).Value
            
            '' set next row for invoice transaction details
            r = r + 1
        
        End If
        End With
        
    Next i
    

End Sub

Sub exportToPdf()

    ''ensure data detail is saved to invoice database
    Call RecordInvoice

    ''ensure invoice data is  the one selected
    Call LoadInvoice
    
    ''Copy selected invoice data to template
    Call TransferInvoiceToTemplate

    Dim invNo As Long
    Dim custName As String
    Dim dateIssued As Date
    
    invNo = Sheet1.Range("J7").Value
    custName = Sheet1.Range("E12").Value
    dateIssued = Sheet1.Range("J9").Value
    
    Dim savePath As String
    Dim fileName As String
    
    ''save pdf to allocated path
    savePath = "C:\Users\ahmad\Documents\Aku\Project\Excel-VBA-Automated-Invoicing-System\invoice\"
    fileName = "INV-" & Year(dateIssued) & "-" & custName & "-" & invNo
    Sheet1.ExportAsFixedFormat xlTypePDF, IgnorePrintAreas:=False, fileName:=savePath & fileName
    
    
    '' to save the link to invoice database(from previous test) Sheet3.Hyperlinks.Add Anchor:=Sheet3.Cells(location), Address:=savePath & fileName & ".pdf"
        ''WIP
End Sub
