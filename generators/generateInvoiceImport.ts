// import { extractColumnsByHeader, getColumnIndexes } from '../Code'

const generateInvoiceImport = picklistData => {
  let desiredHeaders = [
    'Product Code',
    'Buyer Store No',
    'sku',
    'PO',
    'Ship Not Before', 
    'Cancel After',
    'Unit Price',
    'In Stock'
  ] // I can easily read these from somewhere else i.e. input box, sidebar, another sheet, etc.
  let newData = extractColumnsByHeader(picklistData, desiredHeaders)
  // remove the 0 quantity rows
  // get columns for instock, upc, store
  let newHeader = newData[0]
  let { inStockColumnIndex, barcodeColumnIndex, storeColumnIndex } = getColumnIndexes(newHeader)
  newData = newData.filter((row, i) => i === 0 || row[inStockColumnIndex] > 0)
  // sort by upc then store
  newData = newData.sort((a, b) => a[barcodeColumnIndex] - b[barcodeColumnIndex])
  .sort((a, b) => a[storeColumnIndex] - b[storeColumnIndex])
  return newData
}
