# Voneddie

This project aims to automate the process for receiving and fulfilling an EDI order from Von Maur at Marc Joseph New York.

It is written in Google Apps Script as a Google Sheet add-on utilizing [clasp](https://github.com/google/clasp).

There is also an excel `.bas` file that was written to automate excel workflow.

## Steps of the process

Prior to writing the script the following steps had to be taken manually:
  1. Go to [diCentral Portal](https://diwebc.dicentral.com/Main.aspx) and find desired PO number
  2. Select all stores and export as csv
  3. Save as {po#}.xls
  4. Add column for sku via a lookup in a barcode reference on another spreadsheet [AUTO]
  5. Add column for po+store by combining po column and store column [AUTO]
  6. Remove extraneous data, add column called "in stock"  [AUTO]
  7. Sort by upc
  8. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist
  9. Add stock data in column called "in stock",copy values to new sheet called invoiced
  10. sort by in stock, Delete all rows with 0 qty in stock
  11. Sort by upc, then by store #
  12. import via Zed Axis as invoice
  13. Create Pivot table on new sheet with weight calculations - store # (NOT PO) copy pivot table as values then add =ROUNDUP(E4*1.2+1, 0)and add invoice numbers before store # column
  14. Create shipping labels or truck routing
  15. Tracking numbers should be in order of invoices and sent via slack, add tracking # to qb invoice and to asn as well as weight from pivot table, items for stock report, and invoice number from quickbooks (Possibly create another sheet for this)
  16. Print packing slip and ucc, they should be aligned for warehouse
  17. Create EDI invoice based on tracking # from slack, invoice from quickbooks, and remove missing item via warehouse stock report

## TODO

- [x] Get the sku data and insert before first column
- [x] Sort by UPC
- [x] Create another sheet called picklist that has columns sku, upc, po, qty, in stock
- [ ] Send Print job to the printer
- [x] Add another drop down option for after stock was taken
- [x] Generate another sheet that contains invoice data for quickbooks without 0 qty items
- [x] sort by upc, then store
- [x] Create a new sheet for shipping data
  - [x] by store
  - [x] sum units
  - [x] weight calculation =ROUNDUP(E4*1.2+1, 0)
- [ ] Add message box informing that shipping data is available in the shipping tab and that tracking and invoices need to be entered there
- [ ] Add another sheet with shipping data and sku qtys by store
- [ ] Generate another sheet with a checklist of all the steps in each row and a checkbox next to it

