// import { extractColumnsByHeader, getColumnIndexes } from '../Code'

const generateInvoiceImport = picklistData => {
  // remove the 0 quantity rows
  // get columns for instock, upc, store
  let headerRow = picklistData[0]
  let { inStockColumnIndex, barcodeColumnIndex, storeColumnIndex } = getColumnIndexes(headerRow)
  let newData = picklistData.filter((row, i) => i === 0 || row[inStockColumnIndex] > 0)
  // sort by upc then store
  newData = sortByUpc(newData)
  .sort((a, b) => a[storeColumnIndex] - b[storeColumnIndex])
  return newData
}
