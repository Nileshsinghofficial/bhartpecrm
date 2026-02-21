const express = require("express");
const multer = require("multer");
const { chromium } = require("playwright");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

let browser = null;
let context = null;
let page = null;

global.progress = 0;

app.use(express.urlencoded({ extended: true }));

// ================= HOME =================
app.get("/", (req, res) => {
  res.send(`
    <h2>CRM Login</h2>
    <form action="/open-crm" method="post">
      <button type="submit">Open CRM</button>
    </form>
  `);
});

// ================= OPEN CRM =================
app.post("/open-crm", async (req, res) => {

  if (!browser) {
    browser = await chromium.launch({ headless: false });
    context = await browser.newContext();
    page = await context.newPage();
  }

  await page.goto("https://support.bharatpe.in/crm/login");

  res.send(`
    <h3>CRM Opened</h3>
    <p>Email + OTP enter karo</p>
    <a href="/upload-page">Login Complete? Continue</a>
  `);
});

// ================= UPLOAD PAGE =================
app.get("/upload-page", async (req, res) => {

  if (!page || !page.url().includes("/crm/")) {
    return res.send("Login complete nahi hua.");
  }

  res.send(`
    <h2>Upload Excel</h2>
    <form action="/preview" method="post" enctype="multipart/form-data">
      <input type="file" name="file" required />
      <button type="submit">Upload</button>
    </form>

    <br>
    <div style="width:400px;border:1px solid #000;">
      <div id="bar" style="height:20px;width:0%;background:green;"></div>
    </div>

    <script>
      function checkProgress(){
        fetch('/progress')
        .then(r=>r.json())
        .then(d=>{
          document.getElementById('bar').style.width = d.progress + "%";
          if(d.progress < 100){
            setTimeout(checkProgress,1000);
          }
        });
      }
      checkProgress();
    </script>

    <br><a href="/logout">Logout</a>
  `);
});

// ================= PREVIEW SHEETS =================
app.post("/preview", upload.single("file"), (req, res) => {

  const originalName = req.file.originalname;
  const ext = path.extname(originalName); // .xlsx or .csv

  const newPath = path.join("uploads", Date.now() + ext);

  fs.renameSync(req.file.path, newPath);

  const workbook = XLSX.readFile(newPath);
  const sheets = workbook.SheetNames;

  const options = sheets.map(s => `<option value="${s}">${s}</option>`).join("");

  res.send(`
    <h3>Select Sheet</h3>
    <form action="/process" method="post">
      <input type="hidden" name="filePath" value="${newPath}" />
      <select name="sheetName">${options}</select>
      <button type="submit">Start Processing</button>
    </form>
  `);
});

// ================= PROCESS =================
app.post("/process", async (req, res) => {
    if (!page) {
  return res.send("CRM session not active. Please login again.");
}

  const filePath = req.body.filePath;
  const sheetName = req.body.sheetName;

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  global.progress = 0;

  for (let i = 0; i < data.length; i++) {

  let mid = data[i]["MID"];

  // Skip invalid MID
  if (!mid || mid === "-" || isNaN(mid)) {
    continue;
  }

  // 🔥 RESUME CHECK
  if (data[i]["Loan Stage"] && data[i]["Loan Stage"] !== "") {
    console.log("Already processed MID:", mid);
    continue;
  }

  console.log("Processing MID:", mid);

  try {

    await page.goto(
      `https://support.bharatpe.in/crm/merchant-details?mid=${mid}`,
      { waitUntil: "domcontentloaded", timeout: 10000 }
    );

    if (page.url().includes("login")) {
      console.log("Session expired");
      break;
    }

    await page.waitForSelector("text=LOAN", { timeout: 7000 });

    const responsePromise = page.waitForResponse(
      r => r.url().includes("all-details"),
      { timeout: 7000 }
    );

    await page.locator("text=LOAN").first().click();

    const response = await responsePromise;
    const json = await response.json();

    let stage = "";
    let status = "";
    let message = "";

    if (json?.data) {
      const d = json.data;
      stage = d.supportApiResponseDto?.applicationStage || "";
      status = d.applicationStatus || "";
      message =
        d.conditionalMessage ||
        d.stageCommunication ||
        d.message ||
        "";
    }

    data[i]["Loan Stage"] = stage;
    data[i]["Application Status"] = status;
    data[i]["Message"] = message;

  } catch (err) {

    console.log("Error for MID:", mid);

    data[i]["Loan Stage"] = "Error";
    data[i]["Application Status"] = "";
    data[i]["Message"] = "";
  }

  global.progress = Math.floor(((i + 1) / data.length) * 100);
}

  // Append columns safely at end
const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
let headers = sheetData[0];

headers.push("Loan Stage", "Application Status", "Message");

for (let i = 1; i < sheetData.length; i++) {
  sheetData[i].push(data[i - 1]?.["Loan Stage"] || "");
  sheetData[i].push(data[i - 1]?.["Application Status"] || "");
  sheetData[i].push(data[i - 1]?.["Message"] || "");
}

const newSheet = XLSX.utils.aoa_to_sheet(sheetData);
workbook.Sheets[sheetName] = newSheet;
XLSX.writeFile(workbook, filePath);

global.progress = 100;

res.redirect(`/download?file=${filePath}`);

});

// ================= DOWNLOAD =================
app.get("/download", (req, res) => {

  const file = req.query.file;

  const fileName = path.basename(file);

  res.download(file, "processed_" + fileName);

  setTimeout(() => {
    fs.unlink(file, () => {
      console.log("File deleted:", file);
    });
  }, 60000);
});

// ================= PROGRESS =================
app.get("/progress", (req, res) => {
  res.json({ progress: global.progress || 0 });
});

// ================= LOGOUT =================
app.get("/logout", async (req, res) => {

  if (context) await context.close();
  if (browser) await browser.close();

  browser = null;
  context = null;
  page = null;

  global.progress = 0;

  res.send("Logged out. <a href='/'>Login Again</a>");
});

// ================= START =================
const PORT = process.env.PORT || 5000
app.listen(PORT, () => {
  console.log("Server running at http://localhost:5000");
});