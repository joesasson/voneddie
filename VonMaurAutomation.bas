Attribute VB_Name = "VonMaurAutomation"
' Process Steps
' 1. Download CSV and save as {PO number}.xlsx
' 2. Add column for concatenated PO+store - PreparePickList [AUTO]
' 3. Retrieve Sku data and add as new column - PreparePickList (In Data Tab select fix broken links and open source and close it to get sku) [AUTO]
' 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
' 5. Sort by upc [AUTO]
' 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist
' 7. Add stock data in column called "in stock",copy values to new sheet called invoiced
' 8. sort by in stock, Delete all rows with 0 qty in stock
' 9. Sort by upc, then by store #
' 10. import via Zed Axis as invoice
' 11. Create Pivot table on new sheet with weight calculations  - store # (NOT PO) copy pivot table as values then add =ROUNDUP(E4*1.2+1, 0)and add invoice numbers before store # column
' 12. Create shipping labels or truck routing
' 13. Tracking numbers should be in order of invoices and put into shipping sheet, add tracking # to qb invoice and to asn as well as weight from pivot table, items for stock report, and invoice number from quickbooks (Possibly create another sheet for this)
' 14. Print packing slip and ucc, they should be aligned for warehouse
' 15. Create EDI invoice based on tracking # from slack, invoice from quickbooks, and remove missing item via warehouse stock report

Sub PreparePickList()
Attribute PreparePickList.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' PreparePickList Macro
'
  Dim lastrow As Long
  lastrow = Range("D65000").End(xlUp).Row
  
  Columns("EO:EO").Select
  Selection.NumberFormat = "# ?/?"
  
  Columns("A:A").Select
  Selection.Insert Shift:=xlToRight
  Range("A1").Value = "sku"
  Range("A2").Formula = "=VLOOKUP(EP2, 'C:\Users\ADELE PINDEK\Documents\One Drive\OneDrive\Zed Axis\barcodes.csv'!$A:$B, 2, FALSE)"
  Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
  Columns("A:A").Select
  Selection.Insert Shift:=xlToRight
  Range("A1").Value = "PO"
  Range("A2").Formula = "=F2&" + """-""" + "&GB2"
  Range("A2").AutoFill Destination:=Range("A2:A" & lastrow)
  
  DeleteUserDefinedColumns
  Range("AI1").Value = "in stock"
  SortByUPC
  PrintForWarehouse
End Sub

Sub DeleteUserDefinedColumns()
  Range("GI1:JU1").Clear
  ' loop through header row
  ' if cell contains "userDefined", delete entire column
  For I = ActiveSheet.Columns.Count To 1 Step -1
    If InStr(1, Cells(1, I), "UserDefined") Then Columns(I).EntireColumn.Delete
  Next I
End Sub

Sub SortByUPC()
  ' Get Entire Data Range
  Dim lastrow As Long
  lastrow = Range("A65000").End(xlUp).Row
  ' Sort by UPC Column (Y)
  ' AI is HARDCODED
  Range("A1:AI" & lastrow).Sort key1:=Range("Y2"), _
  order1:=xlAscending, Header:=xlYes

    
End Sub

Sub PrintForWarehouse()
'  Hide all columns except for po sku qtyordered and in stock
End Sub

Sub FormatForQBImport()
  
End Sub

Sub CreateShippingDetails()
  InsertPivotTable
End Sub

Sub AddTrackingNumbers()

End Sub

Sub FormatForASNandInv()

End Sub


