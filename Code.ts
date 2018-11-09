function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Open Voneddie Sidebar', 'showInstructions')
    .addItem("Process EDI", "processEdi")
    .addToUi()
}

function showInstructions(){
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Process Steps for Von Maur EDI')
      .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(html)
}

function processEdi(){
  // Process Steps
  // 1. Download CSV and save as {PO number}.xlsx
  // 2. Add column for concatenated PO+store - PreparePickList [AUTO]
  // 3. Retrieve Sku data and add as new column - PreparePickList (In Data Tab select fix broken links and open source and close it to get sku) [AUTO]
  // 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
  // 5. Sort by upc [AUTO]
  // 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist
  // 7. Add stock data in column called "in stock",copy values to new sheet called invoiced
  // 8. sort by in stock, Delete all rows with 0 qty in stock
  // 9. Sort by upc, then by store #
  // 10. import via Zed Axis as invoice
  // 11. Create Pivot table on new sheet with weight calculations  - store # (NOT PO) copy pivot table as values then add =ROUNDUP(E4*1.2+1, 0)and add invoice numbers before store # column
  // 12. Create shipping labels or truck routing
  // 13. Tracking numbers should be in order of invoices and sent via slack, add tracking # to qb invoice and to asn as well as weight from pivot table, items for stock report, and invoice number from quickbooks (Possibly create another sheet for this)
  // 14. Print packing slip and ucc, they should be aligned for warehouse
  // 15. Create EDI invoice based on tracking # from slack, invoice from quickbooks, and remove missing item via warehouse stock report
  let ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1a78mv6dg9-fSPa40VpiARr3Jjcd2amltNmUbO0FkBzY/edit?addon_dry_run=AAnXSK-bLW7mohOE2aG-EtDuUwWEMgh-2eSrgAwnEgBi4qzkf3e3kWwehTjehtB7zZiZqWPWaqYwxlGM8yzcnxl8J46pgT8RJoRteiyI0ncTrP8WehZqUe0JXH3o2DQq1hJyuFUh3JLa#gid=912552240")
  let sheet = ss.getSheets()[0]
  let sheetData = sheet.getDataRange().getValues()
  let prePicklistData = createPrePicklist(sheetData)
  let picklistSheet = createNewSheetWithData(ss, prePicklistData, "pre-picklist")
}

function createPrePicklist(sheetData: Object[][]) {
  // Remove top row
  let newData = sheetData.filter((row, i) => i > 0)
  let headerRow = newData[0]
  // 2. Add column for concatenated PO+store - PreparePickList [AUTO]
  // 4. Remove extraneous data - DeleteUserDefinedColumns, add column called "in stock" - PreparePickList [AUTO]
  newData= deleteUserDefinedColumns(newData, headerRow)
  headerRow = newData[0]
  newData = newData.map((row, i) => {
    if(i === 0){ return ["PO", ...row] }
    let { poColumnIndex, storeColumnIndex } = getColumnIndexes(headerRow)
    let po = row[poColumnIndex]
    let store = row[storeColumnIndex]
    let poWithStore = `${po}-${store}`
    return [poWithStore, ...row]
  })
  // 3. Retrieve Sku data and add as new column - PreparePickList (In Data Tab select fix broken links and open source and close it to get sku) [AUTO]
  // 5. Sort by upc [AUTO]
  // 6. Print for warehouse - hide all columns beside sku, po, qty, and in stock - create new sheet called picklist
  return newData
}

const getColumnIndexes = (headerRow: Array<Object>) => {
  let poColumnIndex = getColumnIndex(headerRow, "PO Number")
  Logger.log({poColumnIndex, headerRow })
  let storeColumnIndex = getColumnIndex(headerRow, "Buyer Store No")
  return {
    poColumnIndex,
    storeColumnIndex
  }
}

const getColumnIndex = (headerRow: Array<Object>, headerTitle) => headerRow.indexOf(headerTitle)

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

