// Process Steps
  // 1. Download CSV and save as {PO number}.xlsx [MANUAL]
  // 2. Add column for concatenated PO+store - PreparePickList [AUTO]
  // 3. Retrieve Sku data and add as new column - PreparePickList (In Data Tab select fix broken links and open source and close it to get sku) [AUTO]
  // 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
  // 5. Sort by upc [AUTO]
  // 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist [AUTO]
  // 7. Add stock data in column called "in stock",copy values to new sheet called invoiced [MANUAL]
  // 8. sort by in stock, Delete all rows with 0 qty in stock
  // 9. Sort by upc, then by store #
  // 10. import via Zed Axis as invoice
  // 11. Create Pivot table on new sheet with weight calculations  - store # (NOT PO) copy pivot table as values then add =ROUNDUP(E4*1.2+1, 0)and add invoice numbers before store # column
  // 12. Create shipping labels or truck routing
  // 13. Tracking numbers should be in order of invoices and sent via slack, add tracking # to qb invoice and to asn as well as weight from pivot table, items for stock report, and invoice number from quickbooks (Possibly create another sheet for this)
  // 14. Print packing slip and ucc, they should be aligned for warehouse
  // 15. Create EDI invoice based on tracking # from slack, invoice from quickbooks, and remove missing item via warehouse stock report

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem("Create Picklist", "createPicklist")
    .addItem("Create Invoice Import and Shipping Calculator", "createInvoiceImport")
    .addItem('Open Voneddie Sidebar', 'showInstructions')
    .addToUi()
}

function showInstructions(){
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Process Steps for Von Maur EDI')
      .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(html)
}

function createInvoiceImport(){
  // 8. sort by in stock, Delete all rows with 0 qty in stock
  // 9. Sort by upc, then by store #
  // 10. import via Zed Axis as invoice
  // 11. Create Pivot table on new sheet with weight calculations  - store # (NOT PO) copy pivot table as values then add =ROUNDUP(E4*1.2+1, 0)and add invoice numbers before store # column
  // 12. Create shipping labels or truck routing
  // 13. Tracking numbers should be in order of invoices and sent via slack, add tracking # to qb invoice and to asn as well as weight from pivot table, items for stock report, and invoice number from quickbooks (Possibly create another sheet for this)
  // 14. Print packing slip and ucc, they should be aligned for warehouse
  // 15. Create EDI invoice based on tracking # from slack, invoice from quickbooks, and remove missing item via warehouse stock report
  let ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1a78mv6dg9-fSPa40VpiARr3Jjcd2amltNmUbO0FkBzY/edit?addon_dry_run=AAnXSK-bLW7mohOE2aG-EtDuUwWEMgh-2eSrgAwnEgBi4qzkf3e3kWwehTjehtB7zZiZqWPWaqYwxlGM8yzcnxl8J46pgT8RJoRteiyI0ncTrP8WehZqUe0JXH3o2DQq1hJyuFUh3JLa#gid=912552240")
  let prePicklist = ss.getSheetByName('pre-picklist')
  let prePicklistData = prePicklist.getDataRange().getValues()
  let invoiceImportData = generateInvoiceImport(prePicklistData)
  createNewSheetWithData(ss, invoiceImportData, "Invoice Import")
  let shippingDetailData = generateShippingDetails(invoiceImportData)
  createNewSheetWithData(ss, shippingDetailData, "Shipping Details")
  let ediDetails = generateEdiDetails()
}

function createPicklist(){
  // Process Steps
  // 1. Download CSV and save as {PO number}.xlsx [MANUAL]
  // 2. Add column for concatenated PO+store - PreparePickList [AUTO]
  // 3. Retrieve Sku data and add as new column - PreparePickList (In Data Tab select fix broken links and open source and close it to get sku) [AUTO]
  // 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
  // 5. Sort by upc [AUTO]
  // 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist [AUTO]
  // 7. Add stock data in column called "in stock",copy values to new sheet called invoiced [MANUAL]
  let ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1a78mv6dg9-fSPa40VpiARr3Jjcd2amltNmUbO0FkBzY/edit?addon_dry_run=AAnXSK-bLW7mohOE2aG-EtDuUwWEMgh-2eSrgAwnEgBi4qzkf3e3kWwehTjehtB7zZiZqWPWaqYwxlGM8yzcnxl8J46pgT8RJoRteiyI0ncTrP8WehZqUe0JXH3o2DQq1hJyuFUh3JLa#gid=912552240")
  let sheet = ss.getSheets()[0]
  let sheetData = sheet.getDataRange().getValues()
  let prePicklistData = generatePrePicklist(sheetData)
  createNewSheetWithData(ss, prePicklistData, "pre-picklist")
  let picklistData = generatePicklist(prePicklistData)
  createNewSheetWithData(ss, picklistData, "picklist")
}

function generatePrePicklist(sheetData: Object[][]) {
  // Remove top row
  let newData = sheetData.filter((row, i) => i > 0)
  let headerRow = newData[0]
  // 2. Add column for concatenated PO+store - PreparePickList [AUTO]
  // 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
  newData = deleteUserDefinedColumns(newData, headerRow)
  headerRow = newData[0]
  let { poColumnIndex, storeColumnIndex, barcodeColumnIndex } = getColumnIndexes(headerRow)
  newData = newData.map((row, i) => {  
    if(i === 0){ return ["sku", "PO", ...row, "In Stock"] }
    // Add sku column
    // 3. Retrieve Sku data and add as new column - PreparePickList
    let sku = getSkuFromBarcodeReference(row[barcodeColumnIndex])
    // add PO column
    let po = row[poColumnIndex]
    let store = row[storeColumnIndex]
    let poWithStore = `${po}-${store}`
    return [sku, poWithStore, ...row, ""]
  })
  // 5. Sort by upc
  newData = sortByUpc(newData)
  // 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist
  return newData
}

const getColumnIndexes = (headerRow: Array<Object>) => {
  let ColumnNames = [
    {
      header: "PO Number",
      columnName: "po"
    },
    {
      header: "Buyer Store No",
      columnName: 'store'
    },
    {
      header: 'Product Code',
      columnName: 'barcode'
    },
    {
      header: "sku",
      columnName: "sku"
    },
    {
      header: "PO",
      columnName: 'newPo'
    },
    {
      header: "In Stock",
      columnName: 'inStock'
    },
    {
      header: "Ship Not Before",
      columnName: 'shipDate'
    },
    {
      header: "Cancel After",
      columnName: 'cancelDate'
    },
    {
      header: "Unit Price",
      columnName: 'price'
    },
  ]
  let indexes = getColumns(headerRow, ColumnNames)
  return indexes
}

const getColumns = (headerRow: Object[], columnNames) => {
  let indexes = {}
  columnNames.forEach(col => indexes[`${col.columnName}ColumnIndex`] = getColumnIndex(headerRow, col.header))
  return indexes
}

const getColumnIndex = (headerRow: Array<Object>, headerTitle: String, name?: String) =>  { 
  return headerRow.indexOf(headerTitle)
}

const deleteUserDefinedColumns = (sheetData: Object[][], headerRow: Array<Object>) => {
  // I want to delete the nth element of each of the rows
  // Find the indexes of all the columns that I want to remove
  let userDefinedIndexes = findUserDefinedIndexes(headerRow)
  // map through each row and return only the elements at the columns that are not in the array of indexes
  let newData = sheetData.map(row => row.filter((col, i) => !(userDefinedIndexes.indexOf(i) > -1)))
  return newData
}

const findUserDefinedIndexes = (headerRow: Array<Object>) => {
  return getAllColumnIndexesByHeader(headerRow, "UserDefined")
}

const getAllColumnIndexesByHeader = (headerRow: Array<Object>, headerTitle: Object) => {
  return headerRow.map((col, i) => {
    if(col === headerTitle) { 
      return i
    }
  }).filter(col => col)
}

const createNewSheetWithData = (ss: GoogleAppsScript.Spreadsheet.Spreadsheet, data, sheetName) => {
  // find if sheetName exists, if so delete it
  let previousSheet = ss.getSheetByName(sheetName)
  let newSheet: GoogleAppsScript.Spreadsheet.Sheet
  if(previousSheet){
    newSheet = previousSheet.clear()
  } else {
    newSheet = ss.insertSheet(sheetName)
  }
  // get dimensions of data
  let dataHeight = data.length
  let dataWidth = data[0].length
  // set data on new sheet based on dimensions of data
  let targetRange = newSheet.getRange(1, 1, dataHeight, dataWidth)
  targetRange.setValues(data)
  return newSheet
}

const getSkuFromBarcodeReference = upc => {
  const url = 'https://sku-barcode-lookup.herokuapp.com/graphql'
  const payload = {
    query:  
    `{ pair(upc:"${upc}") { sku } }` 
  }
  const options = {
    method: "post",
    contentType: 'application/json' ,
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  }
  //@ts-ignore: 
  // Argument of type '{ method: string;}' is not assignable to parameter of type 'URLFetchRequestOptions'.
  // Types of property 'method' are incompatible.
  // Type 'string' is not assignable to type '"post" | "get" | "delete" | "patch" | "put"'. 
  let response = UrlFetchApp.fetch(url, options).getContentText()
  let skus = JSON.parse(response).data.pair.sku
  return skus
}

const sortByUpc = (data) => {
  let headerRow = data[0]
  let { barcodeColumnIndex } = getColumnIndexes(headerRow)
  let newData = data.sort((a, b) => Number(a[barcodeColumnIndex]) - Number(b[barcodeColumnIndex]))
  return newData
}

const generatePicklist = data => {
  let headerRow = data[0]
  // map through returning only the columns that I want
  // Refactor to include only [sku, po, qty, and in stock]
  let columnIndices: Object = getColumnIndexes(headerRow)
  let headerIndices = Object.keys(columnIndices).map(key => columnIndices[key]).sort((a, b) => a - b) // [0, 1, 4, 24]
  return data.map(row => {
    let newRow = []
    headerIndices.forEach(i => { newRow.push(row[i])})
    return newRow
  })
}

const extractColumnsByHeader = (sheetData: Object[][], desiredHeaders: String[]) => {
  let headerRow = sheetData[0]
  // map headers into indexes
  let indices = desiredHeaders.map(header => headerRow.indexOf(header)).filter(x => x === 0 || x)
  // map through each row and return only if column index is in indices
  let newData = sheetData.map(row => {
    return row.map((el, i) => {
      if(indices.indexOf(i) > -1){
        return el
      }
    }).filter(x => x === 0 || x)
  })
  return newData
}

const generateShippingDetails = invoiceData => {
  // pivot data to store number, sum of in stock qty, weight calculation
  // the rest is manual invoice and tracking number after the import
  // create an object { storenumber1: sumqty, storenumber2: sumqty }
  let sumsByStore = sumStoreQtys(invoiceData)
  Logger.log(sumsByStore)
  // then map through keys and return [storenumber, qty, weight,'','']
  // headers should be ['storenumber', 'qty', 'weight', 'invoice', 'tracking #']
  let shippingDetails = getShippingDetails(sumsByStore)
  Logger.log(shippingDetails)
  return shippingDetails
}

const generateEdiDetails = () => {}

const sumStoreQtys = sheetData => {
  let headerRow = sheetData[0]
  let { storeColumnIndex, inStockColumnIndex } = getColumnIndexes(headerRow)
  let qtysByStore = {}
  sheetData.forEach(row => {
    let store = row[storeColumnIndex]
    let qty = row[inStockColumnIndex]
    let currentQty = qtysByStore[store]
    currentQty ? qtysByStore[store] = currentQty + qty : qtysByStore[store] = qty
  })
  return qtysByStore
}

const getShippingDetails = storeQtys => {
  return Object.keys(storeQtys).map((key, i) => {
    if(isNaN(Number(key))){
      return [key, storeQtys[key], 'Weight', 'Invoice #', 'Tracking #']
    }
    let qty = storeQtys[key]
    let weight = calculateWeight(qty)
    return [key, qty, weight, '', '']
  })
}

const calculateWeight = qty => Math.ceil(qty * 1.2 + 1)