import './App.css'
import * as XLSX from 'xlsx'
import { useState } from 'react'
import 'bootstrap/dist/css/bootstrap.min.css'
import { Container, Form, Button, Table } from 'react-bootstrap'
import { saveAs } from 'file-saver'

function App() {
  const [openingStockFile, setOpeningStockFile] = useState()
  const [purchaseFile, setPurchaseFile] = useState()
  const [saleFile, setSaleFile] = useState()
  const [closingStockReport, setClosingStockReport] = useState([])

  function readFileAsync(file) {
    return new Promise((resolve, reject) => {
      let reader = new FileReader()
      reader.onload = () => {
        resolve(reader.result)
      }
      reader.onerror = reject
      reader.readAsArrayBuffer(file)
    })
  }

  const handleFiles = async () => {
    const openingStockReport = []
    const purchaseReport = []
    const purchaseReportSorted = []
    const saleReport = []
    const saleReportSorted = []

    // Opening Stock Report
    let ab = await readFileAsync(openingStockFile)
    let wss = XLSX.read(ab, { type: 'array' })
    let wsname = wss.SheetNames[0]
    let ws = wss.Sheets[wsname]
    let data = XLSX.utils.sheet_to_json(ws)

    data.forEach((item) => {
      openingStockReport.push({ name: item['Item Name'], qty: +item['Stock Quantity'] })
    })

    // Purchase Report
    ab = await readFileAsync(purchaseFile)
    wss = XLSX.read(ab, { type: 'array' })
    wsname = wss.SheetNames[1]
    ws = wss.Sheets[wsname]
    data = XLSX.utils.sheet_to_json(ws)

    data.forEach((item) => {
      purchaseReport.push({ name: item['Item Name'], qty: item['Quantity'] })
    })

    purchaseReport.forEach(function (a) {
      if (!this[a.name]) {
        this[a.name] = { name: a.name, qty: 0 }
        purchaseReportSorted.push(this[a.name])
      }
      this[a.name].qty += +a.qty
    }, Object.create(null))

    // Sale Report
    ab = await readFileAsync(saleFile)
    wss = XLSX.read(ab, { type: 'array' })
    wsname = wss.SheetNames[1]
    ws = wss.Sheets[wsname]
    data = XLSX.utils.sheet_to_json(ws)

    data.forEach((item) => {
      saleReport.push({ name: item['Item Name'], qty: item['Quantity'] })
    })

    saleReport.forEach(function (a) {
      if (!this[a.name]) {
        this[a.name] = { name: a.name, qty: 0 }
        saleReportSorted.push(this[a.name])
      }
      this[a.name].qty += +a.qty
    }, Object.create(null))

    // Closing Stock Report
    const _closingStockReport = []
    openingStockReport.forEach((item) => {
      const csrItem = item
      purchaseReportSorted.forEach((el) => {
        if (el.name !== item.name) return
        csrItem.qty += el.qty
      })
      saleReportSorted.forEach((el) => {
        if (el.name !== item.name) return
        csrItem.qty -= el.qty
      })
      _closingStockReport.push(csrItem)
    })

    setClosingStockReport(_closingStockReport)
  }

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(closingStockReport)
    const workbook = {
      Sheets: {
        sheet_1: ws,
      },
      SheetNames: ['sheet_1'],
    }
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })

    const data = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    })

    saveAs(data, 'closing_stock_report.xlsx')
  }

  return (
    <div className="App">
      <Container>
        <h1 className="my-4 fw-bold text-center">Stockify</h1>

        <Form>
          <Form.Group className="mb-3">
            <Form.Label>Opening Stock Report</Form.Label>
            <Form.Control
              type="file"
              onChange={(event) => setOpeningStockFile(event.target.files[0])}
            />
          </Form.Group>
          <Form.Group className="mb-3">
            <Form.Label>Purchase Report</Form.Label>
            <Form.Control
              type="file"
              onChange={(event) => setPurchaseFile(event.target.files[0])}
            />
          </Form.Group>
          <Form.Group className="mb-3">
            <Form.Label>Sale Report</Form.Label>
            <Form.Control type="file" onChange={(event) => setSaleFile(event.target.files[0])} />
          </Form.Group>
          <Button variant="primary" type="button" onClick={handleFiles}>
            Run
          </Button>
        </Form>

        <div className="d-flex justify-content-between align-items-center mb-2 mt-4">
          <h4>Closing Stock Report</h4>
          <Button onClick={exportToExcel}>Export as Excel</Button>
        </div>
        <Table>
          <thead>
            <tr>
              <th>#</th>
              <th>Item Name</th>
              <th>Quantity</th>
            </tr>
          </thead>
          <tbody>
            {closingStockReport.map((item, index) => (
              <tr key={item.name}>
                <td>{index + 1}</td>
                <td>{item.name}</td>
                <td>{item.qty}</td>
              </tr>
            ))}
          </tbody>
        </Table>
      </Container>
    </div>
  )
}

export default App
