import express from "express";
import bodyParser from "body-parser";
import cors from "cors";
import XLSX from "xlsx";
import path from "path";
import dotenv from "dotenv";
import fs from "fs";

dotenv.config(); // Load .env variables

const app = express();

// Use PORT from environment or fallback for local testing
const PORT = process.env.PORT || 5000;

// Use Excel file path from .env or fallback
const excelFilePath = path.join(process.cwd(), process.env.EXCEL_FILE || "stocks.xlsx");

app.use(cors());
app.use(bodyParser.json());

// Load data from Excel
const loadData = () => {
  if (!fs.existsSync(excelFilePath)) return [];
  const wb = XLSX.readFile(excelFilePath);
  const ws = wb.Sheets["Sheet1"];
  return XLSX.utils.sheet_to_json(ws);
};

// Save data to Excel
const saveData = (data) => {
  let wb;
  if (fs.existsSync(excelFilePath)) {
    wb = XLSX.readFile(excelFilePath);
  } else {
    wb = XLSX.utils.book_new();
  }
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, excelFilePath);
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

// Debug route: check headers in Excel
app.get("/debug-headers", (req, res) => {
  const data = loadData();

  if (data.length === 0) {
    return res.json({ message: "No data found in Excel" });
  }

  res.json({
    headers: Object.keys(data[0]),
    firstRow: data[0],
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running at port ${PORT}`);
});
