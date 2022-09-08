const ExcelJS = require('exceljs')
const origins = require('./origins')
const products = require('./products')
const sellers = require('./sellers')

const workbook = new ExcelJS.Workbook()

workbook.creator = 'Random'
workbook.created = new Date()
workbook.lastModifiedBy = new Date()

const mainSheet = workbook.addWorksheet('Leads', { state: 'visible' })
const originsSheet = workbook.addWorksheet('Origins', { state: 'veryHidden' })
const productsSheet = workbook.addWorksheet('Products', { state: 'veryHidden' })
const sellersSheet = workbook.addWorksheet('Users', { state: 'veryHidden' })

originsSheet.columns = [
  { header: 'ID', key: 'originId' },
  { header: 'Origem', key: 'originName' }
]

productsSheet.columns = [
  { header: 'ID', key: 'productId' },
  { header: 'Produto', key: 'productName' }
]

sellersSheet.columns = [
  { header: 'ID', key: 'sellerId' },
  { header: 'Vendedor', key: 'sellerName' }
]

originsSheet.addRows(
  origins.map((origin) => ({ originId: origin.id, originName: origin.name }))
)

productsSheet.addRows(
  products.map((product) => ({
    productId: product.id,
    productName: product.name
  }))
)

sellersSheet.addRows(
  sellers.map((seller) => ({
    sellerId: seller.id,
    sellerName: seller.name
  }))
)

mainSheet.columns = [
  { header: 'Nome', key: 'name' },
  { header: 'E-mail', key: 'email' },
  { header: 'Origem', key: 'originName' },
  { header: 'Código origem', key: 'originId', hidden: false },
  { header: 'Produto', key: 'producName' },
  { header: 'Código Produto', key: 'productId', hidden: false },
  { header: 'Código Unidade de negócio', key: 'businessUnitId', hidden: false },
  { header: 'Código Ponto de Venda', key: 'salePointId', hidden: false },
  { header: 'Vendedor', key: 'sellerName' },
  { header: 'Código Vendedor', key: 'sellerId', hidden: false },
  { header: 'Data de Criação', key: 'created_at' }
]

mainSheet.addRow(null)
mainSheet.duplicateRow(2, 998, false)

const nameColsOrigins = mainSheet.getColumnKey('originName')
const nameColsProducts = mainSheet.getColumnKey('producName')
const nameColsSellers = mainSheet.getColumnKey('sellerName')
const nameColsSellerPoint = mainSheet.getColumnKey('salePointId')
const nameColsSellerBusinessUnit = mainSheet.getColumnKey('businessUnitId')

nameColsSellerPoint.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
  if (rowNumber === 1) return

  cell.value = 134343
})

nameColsSellerBusinessUnit.eachCell(
  { includeEmpty: true },
  (cell, rowNumber) => {
    if (rowNumber === 1) return

    cell.value = 38398283
  }
)

nameColsOrigins.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
  if (rowNumber === 1) return

  const nextCell = mainSheet.getCell(cell.row, cell.col + 1)

  nextCell.value = {
    formula: `INDEX(Origins!A2:B${origins.length + 1},MATCH(${
      cell.address
    },Origins!B2:B${origins.length + 1},0),1)`
  }

  cell.dataValidation = {
    type: 'list',
    errorStyle: 'error',
    showErrorMessage: true,
    allowBlank: false,
    formulae: [`"${origins.map((origin) => origin.name).join(',')}"`]
  }
})

nameColsProducts.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
  if (rowNumber === 1) return

  const nextCell = mainSheet.getCell(cell.row, cell.col + 1)

  nextCell.value = {
    formula: `INDEX(Products!A2:B${products.length + 1},MATCH(${
      cell.address
    },Products!B2:B${products.length + 1},0),1)`
  }

  cell.dataValidation = {
    type: 'list',
    errorStyle: 'error',
    showErrorMessage: true,
    allowBlank: false,
    formulae: [`"${products.map((product) => product.name).join(',')}"`]
  }
})

nameColsSellers.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
  if (rowNumber === 1) return

  const nextCell = mainSheet.getCell(cell.row, cell.col + 1)

  nextCell.value = {
    formula: `INDEX(Products!A2:B${sellers.length + 1},MATCH(${
      cell.address
    },Users!B2:B${sellers.length + 1},0),1)`
  }

  cell.dataValidation = {
    type: 'list',
    errorStyle: 'error',
    showErrorMessage: true,
    allowBlank: false,
    formulae: [`"${sellers.map((seller) => seller.name).join(',')}"`]
  }
})

workbook.xlsx.writeFile('leads.xlsx')
