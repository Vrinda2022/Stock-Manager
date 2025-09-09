const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const XLSX = require("xlsx");
const fs = require("fs");

const app = express();
const PORT = 5000;

app.use(cors());
app.use(bodyParser.json());

const filePath = "./stocks.xlsx";

const loadData = () => {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets["Sheet1"];
  return XLSX.utils.sheet_to_json(ws);
};

const saveData = (data) => {
    const wb = XLSX.readFile(filePath);
    const ws = XLSX.utils.json_to_sheet(data);
    wb.Sheets["Sheet1"] = ws;
    XLSX.writeFile(wb, filePath);   
};

// Update stock by product code
app.post("/update-stock", (req, res) => {
  const { code, retail, billing, remarks } = req.body;
  let data = loadData();

  let item = data.find((d) => String(d.Code) === String(code));

  if (item) {
    if (retail !== undefined) item.Retail = Number(retail);
    if (billing !== undefined) item.Billing = Number(billing);
    if (remarks !== undefined) item.Remarks = remarks;
  } else {
    return res.status(404).json({ error: "Product not found" });
  }

  saveData(data);
  res.json({ success: true, message: "Stock updated!", item });
});

// Get all stocks
app.get("/stocks", (req, res) => {
  res.json(loadData());
});

// Analysis: total available, sold, low stock, high stock
app.get("/analysis", (req, res) => {
  const data = loadData();
  if (data.length === 0) return res.json({ message: "No data available" });

  const totalStock = data.reduce((sum, i) => sum + (Number(i.Retail) || 0), 0);
  const totalSold = data.reduce((sum, i) => sum + (Number(i.Billing) || 0), 0);

  const lowStock = data.reduce((min, i) =>
    (Number(i.Retail) || 0) < (Number(min.Retail) || Infinity) ? i : min
  );
  const highStock = data.reduce((max, i) =>
    (Number(i.Retail) || 0) > (Number(max.Retail) || 0) ? i : max
  );

  res.json({
    totalAvailable: totalStock,
    totalSold: totalSold,
    lowestStock: lowStock,
    highestStock: highStock,
  });
});

// Debug route: check the headers in the Excel
app.get("/debug-headers", (req, res) => {
  const data = loadData();

  if (data.length === 0) {
    return res.json({ message: "No data found in Excel" });
  }

  // Return the first row to see the keys
  res.json({
    headers: Object.keys(data[0]),
    firstRow: data[0],
  });
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
