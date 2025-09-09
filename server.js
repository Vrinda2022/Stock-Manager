const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(bodyParser.json());

// Excel file path
const filePath = path.join(__dirname, "stocks.xlsx");

// Serve index.html (frontend) at root
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

// -------- Excel API --------
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

// Update stock
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

// Analysis
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

// ✅ Express v5 compatible fallback (instead of app.get("*"))
app.use((req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(PORT, () => {
  console.log(`✅ Server running at http://localhost:${PORT}`);
});
