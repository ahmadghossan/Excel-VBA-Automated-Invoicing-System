Attribute VB_Name = "modCreateInvoice"
Option Explicit

Public Sub NewInvoice()

Dim invNo As Long

With Sheet7

invNo = .Range("V8")

Call ClearInvoiceForm

Range("V8") = invNo + 1
Range("V10") = Date


End With

End Sub

Public Sub resetFormula()

    Dim wsInput As Worksheet
    Set wsInput = Sheet7

    Dim i As Long
    
    For i = 33 To 47
    
    With wsInput
    
    .Cells(i, FindHeaderColumn(wsInput, 32, "Description")).Formula = "=IFERROR(VLOOKUP($H" & i & ",'Product Database'!A:G, 2,FALSE),"""")"
    .Cells(i, FindHeaderColumn(wsInput, 32, "Unit")).Formula = "=IFERROR(VLOOKUP($H" & i & ",'Product Database'!A:G, 4,FALSE),"""")"
    .Cells(i, FindHeaderColumn(wsInput, 32, "Price/Unit")).Formula = "=IFERROR(VLOOKUP($H" & i & " , 'Product Database'!A:G, 6, FALSE), """")"
    
    End With
    
    Next i
        

End Sub

Public Sub ClearInvoiceForm()

    With Sheet7
    
    Range("J6:L6").ClearContents
    Range("V8:W11").ClearContents
    Range("K16:L16").ClearContents
    Range("H33:J47").ClearContents
    Range("P33:Q47").ClearContents
    Range("G22:L26").ClearContents
    Range("J6:L6") = "(draft)"
    Range("K16") = "(not selected)"
    
    End With
    

End Sub

Public Sub RecordInvoice()

If Sheet7.Range("selectedCust") = "(not selected)" Or _
    Sheet7.Range("selectedCust") = "" _
Then
    MsgBox ("Client/Customer not selected")
    Exit Sub
    
End If

Dim invNo As Long
Dim poNo As Long
Dim custName As String
Dim amt As Currency
Dim dateIssued As Date
Dim term As Long
Dim nextRow As Long
Dim foundInv As Range

With Sheet7

invNo = .Range("V8")
poNo = .Range("V9")
custName = .Range("K16")
dateIssued = .Range("V10")
term = .Range("V11")

End With

'detecting last row

nextRow = Sheet3.Cells(Sheet3.Rows.Count, "A").End(xlUp).Row + 1

'is invoice already recorded?

Set foundInv = Sheet3.Range("A:A").Find(What:=invNo, LookIn:=xlValues, LookAt:=xlWhole)

'make new row or is invoice already recorded?

If foundInv Is Nothing Then

    nextRow = nextRow
    
Else

    nextRow = foundInv.Row

End If


'record invoice

Sheet7.Range("J6") = "INV-" & Year(dateIssued) & "-" & custName & "-" & invNo

With Sheet3
    
.Cells(nextRow, "A").Value = invNo
    .Cells(nextRow, "B").Value = poNo
    .Cells(nextRow, "C").Value = custName
    .Cells(nextRow, "D").Value = amt
    .Cells(nextRow, "E").Value = dateIssued
    .Cells(nextRow, "F").Value = dateIssued + term
    .Cells(nextRow, "K").Value = "INV-" & Year(dateIssued) & "-" & custName & "-" & invNo

If foundInv Is Nothing Then

    Sheet3.Cells(nextRow, "G").Value = "Unpaid"

End If
    
End With



Call SaveInvoiceDetails


End Sub

Public Sub LoadInvoice()


Dim invName As String
Dim invNo As Long
Dim poNo As Long
Dim custName As String
Dim amt As Currency
Dim dateIssued As Date
Dim term As Long
Dim invRow As Long
Dim foundInv As Range

invName = Sheet7.Range("J6")

'is invname invalid?

If UCase(Left(invName, 3)) <> "INV" Then
    MsgBox "Invalid Invoice", vbQuestion, "Error"
    Exit Sub
End If

'is invoice already recorded?

Set foundInv = Sheet3.Range("K:K").Find(What:=invName, LookIn:=xlValues, LookAt:=xlWhole)

'make new row or is invoice already recorded?

If foundInv Is Nothing Then

MsgBox "Invoice with the name " & invName & " can't be found", vbOKOnly, "Not Found"



    
Else
    
    invRow = foundInv.Row
    
    Call ClearInvoiceForm
    
'record invoice details
With Sheet3
    
    invNo = .Cells(invRow, "A").Value
    poNo = .Cells(invRow, "B").Value
    custName = .Cells(invRow, "C").Value
    dateIssued = .Cells(invRow, "E").Value
    term = .Cells(invRow, "F").Value - dateIssued

End With


'show record invoice details
    
With Sheet7

.Range("V8") = invNo
.Range("V9") = poNo
.Range("K16") = custName
.Range("V10") = dateIssued
.Range("V11") = term
.Range("J6") = invName

End With


End If

Call LoadInvoiceDetails

End Sub


Public Sub SaveInvoiceDetails()
'sheet naming
    Dim wsInput As Worksheet
    Dim wsDB As Worksheet

    Dim tbl As ListObject
    Dim newRow As ListRow

    Dim invoiceName As String

'data column
    Dim skuCol As Long
    Dim descCol As Long
    Dim unitCol As Long
    Dim qtyCol As Long
    Dim priceCol As Long
    Dim amountCol As Long

'for loop
    Dim r As Long 'row to take data

    Set wsInput = Sheet7
    
 'set which sheet and table
 
    Set wsDB = Sheet10
    Set tbl = wsDB.ListObjects("tblInvoiceDetails")

 
    
'Selected invoice
    invoiceName = wsInput.Range("SelectedInvoice").Value

       
'check is invoice details already saved, if found, delete it to be replaced
    
    Dim found As Long
    
    found = FindHeaderRow(Sheet10, "A:A", invoiceName)
    
    If found <> 0 Then
    
    Call DeleteInvoiceDetails
    
    End If
        

'Find columns based on headers
    skuCol = FindHeaderColumn(wsInput, 32, "SKU")
    descCol = FindHeaderColumn(wsInput, 32, "Description")
    unitCol = FindHeaderColumn(wsInput, 32, "Unit")
    qtyCol = FindHeaderColumn(wsInput, 32, "Qty")
    priceCol = FindHeaderColumn(wsInput, 32, "Price/Unit")
    amountCol = FindHeaderColumn(wsInput, 32, "Amount")

'delete data to be replaced if there already data

Call DeleteInvoiceDetails


'take data based on if SKU and Qty is filled for every row on create invoice

    For r = 33 To 47
    
        If wsInput.Cells(r, skuCol) <> "" And _
            wsInput.Cells(r, qtyCol) > 0 Then _
            
            Set newRow = tbl.ListRows.Add
            
            With newRow.Range
            .Cells(1, 1) = invoiceName
            .Cells(1, 2) = wsInput.Cells(r, skuCol)
            .Cells(1, 3) = wsInput.Cells(r, descCol)
            .Cells(1, 4) = wsInput.Cells(r, unitCol)
            .Cells(1, 5) = wsInput.Cells(r, qtyCol)
            .Cells(1, 6) = wsInput.Cells(r, priceCol)
            .Cells(1, 7) = wsInput.Cells(r, amountCol)
            End With
        
        End If
    
    Next r

End Sub

Public Sub LoadInvoiceDetails()

Dim tblRow As Long
Dim trgtRow As Long

Dim tbl As ListObject
Set tbl = Sheet10.ListObjects("tblInvoiceDetails")

'where is invoice column?
Dim invCol As Long
invCol = tbl.ListColumns("Invoice_Name").Index

'invoice details data

Dim skuCol As Long
Dim qtyCol As Long
Dim descCol As Long
Dim unitCol As Long
Dim priceCol As Long

skuCol = FindHeaderColumn(Sheet7, 32, "SKU")
qtyCol = FindHeaderColumn(Sheet7, 32, "Qty")
priceCol = FindHeaderColumn(Sheet7, 32, "Price/Unit")

'which invoice to be search?
Dim invName As String
invName = Sheet7.Range("SelectedInvoice")

'define first row of the item details
trgtRow = 33



'take data from row if inv name the same
    For tblRow = 1 To tbl.ListRows.Count
    
    Debug.Print "Row: " & tblRow
Debug.Print "DB Invoice: [" & _
    tbl.DataBodyRange.Cells(tblRow, invCol).Value & "]"
Debug.Print "Selected: [" & invName & "]"
    
        If tbl.DataBodyRange(tblRow, invCol) = invName Then _
        
            Sheet7.Cells(trgtRow, skuCol) = tbl.DataBodyRange.Cells(tblRow, 2)
            Sheet7.Cells(trgtRow, qtyCol) = tbl.DataBodyRange.Cells(tblRow, 5)
            
        
            trgtRow = trgtRow + 1
            
        End If
        
    
    Next tblRow
    




End Sub

Public Sub DeleteInvoiceDetails()

Dim invCol As Long
Dim i As Long
Dim tbl As ListObject

Dim invName As String

invName = Sheet7.Range("SelectedInvoice").Value

'finding invoice name column

invCol = FindHeaderColumn(Sheet10, 1, "Invoice_Name")

'select tabble

Set tbl = Sheet10.ListObjects("tblInvoiceDetails")

'check ever row on table for last row table to fisrt

    For i = tbl.ListRows.Count To 1 Step -1
    
        'deleting invoice row with the same invoice name
        If tbl.DataBodyRange.Cells(i, invCol) = invName Then _
        
            tbl.ListRows(i).Delete
            
        End If
    
    Next i
    
    
End Sub



