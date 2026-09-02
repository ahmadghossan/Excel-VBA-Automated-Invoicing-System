# 📑 Excel VBA Project: "Automated Invoicing System"

![Project Status](https://img.shields.io/badge/Status-In%20Progress-yellow?style=for-the-badge)
![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Automation-blue?style=for-the-badge)




> 🚧 **Project Status:** This project is currently under active development.  
> Features and source code will continue to be updated as the system is expanded.

## Changelog

### September 2026
- Added **Manage Invoice** page.
  - Added scrollable invoice list.
  - Added Status, Customer filtering.
  - Added helper-range filtering using `FILTER()`, `LET()`, and `CHOOSECOLS()`.


### August 2026
- Added Create Invoice interface.
- Added Save Invoice Details workflow.
- Added Load Invoice Details workflow.
- Added export invoice to PDF feature.

<br>

## 📌 Project Overview & Problem Statement

Managing invoices manually in Excel can become repetitive and increasingly difficult as the number of customers, products, and transactions grows.

Common challenges include:

- Re-entering customer and product information for every invoice.
- Maintaining multiple invoice line items manually.
- Preventing duplicate invoice records.
- Retrieving previously created invoices for review or editing.
- Keeping invoice data structured for future reporting and analysis.
- Maintaining consistency between the invoice form and the underlying transaction database.

This **Excel VBA Automated Invoicing System** is being developed as a practical business automation project using Microsoft Excel and VBA.

The project aims to transform a traditional Excel invoice workbook into a more structured application where users can create, save, retrieve, and eventually update invoices through a user-friendly interface while maintaining invoice records in structured Excel Tables.


> ## Project Frontpage Preview (WIP)
### User form create invoice front page:
![preview create invoice page](/image/CreateInvoiceUserFormPage.png)  


## 🛠️ What I've Learned in this Project So far

### VBA Programming Fundamentals

- Declaring variables using `Dim`.
- Understanding VBA data types such as:
  - `String`
  - `Long`
  - `Double`
  - `Boolean`
  - `Range`
  - `Worksheet`
  - `ListObject`
  - `ListRow`
- Understanding the difference between values and objects.
- Using `Set` when assigning VBA object references.
- Creating reusable `Sub` procedures and `Function` procedures.
- Passing values into procedures using parameters such as `ByVal`.

---

### Control Flow & Data Processing

- Using `For...Next` loops to process multiple invoice line items.
- Using conditional logic with `If...Then`.

Example:

```vb
For i = tbl.ListRows.Count To 1 Step -1

    If tbl.DataBodyRange.Cells(i, invoiceCol).Value = invoiceName Then
        tbl.ListRows(i).Delete
    End If

Next i
```

---

### Excel VBA Object Model

- Working with Excel VBA objects such as:
  - `Workbook`
  - `Worksheet`
  - `Range`
  - `ListObject`
  - `ListRow`
  - `ListColumns`
  - `DataBodyRange`
- Understanding the difference between worksheet-relative and table-relative references.
- Accessing Excel Tables using `ListObjects`.
- Reading and writing data using `Cells(row, column)`.

---

### Excel Formula
  - Using `Let` function
  - `choosecols`
  - `filter`

### Excel Form Control
- Adding interactive elements such as scroll bars to create user friendly environment


