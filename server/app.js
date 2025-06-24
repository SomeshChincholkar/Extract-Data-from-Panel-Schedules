import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import dotenv from "dotenv";
import fetch from "node-fetch";
import path from "path";
import ExcelJS from "exceljs";

dotenv.config();
const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) => cb(null, Date.now() + path.extname(file.originalname)),
});
const upload = multer({ storage });

app.post("/upload", upload.single("image"), async (req, res) => {
  try {
    const imagePath = req.file.path;
    const imageData = await fs.promises.readFile(imagePath);
    const base64Image = imageData.toString("base64");

    // Call Power Automate
    const response = await fetch(process.env.POWER_AUTOMATE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_base64: base64Image }),
    });

    const rawText = await response.text();

    const jsonStart = rawText.indexOf("{");
    const jsonEnd = rawText.lastIndexOf("}");
    if (jsonStart === -1 || jsonEnd === -1) {
      fs.unlinkSync(imagePath);
      return res.status(400).send("JSON not found in response.");
    }

    const jsonString = rawText.substring(jsonStart, jsonEnd + 1);
    const { odd = [], even = [] } = JSON.parse(jsonString);

    // Create Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("PanelBoard");

    worksheet.columns = [
      { header: "Panel No (Odd)", key: "oddNo", width: 18 },
      { header: "Panel Name (Odd)", key: "oddName", width: 40 },
      { header: "Panel No (Even)", key: "evenNo", width: 18 },
      { header: "Panel Name (Even)", key: "evenName", width: 40 },
    ];

    const maxLen = Math.max(odd.length, even.length);
    for (let i = 0; i < maxLen; i++) {
      const oddEntry = odd[i] || {};
      const evenEntry = even[i] || {};
      worksheet.addRow({
        oddNo: oddEntry.panel_no || "",
        oddName: oddEntry.panel_name || "",
        evenNo: evenEntry.panel_no || "",
        evenName: evenEntry.panel_name || "",
      });
    }

    // === IMAGE HANDLING ===
    // Ensure the file is PNG/JPEG only
    const ext = path.extname(imagePath).toLowerCase();
    const allowed = [".png", ".jpg", ".jpeg"];
    if (!allowed.includes(ext)) {
      fs.unlinkSync(imagePath);
      return res.status(400).send("Unsupported image format.");
    }

    // Add space to the right
    worksheet.getColumn(5).width = 50;

    // Add space for image height
    for (let i = 1; i <= 25; i++) {
      worksheet.getRow(i).height = 25;
    }

    // Add image
    const imageId = workbook.addImage({
      buffer: imageData,
      extension: ext.replace(".", ""),
    });

    worksheet.mergeCells("E1:H25");
    worksheet.addImage(imageId, {
      tl: { col: 4.5, row: 0.5 },
      ext: { width: 400, height: 300 },
    });

    // Clean up
    fs.unlinkSync(imagePath);

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=Panel_Board_Listing.xlsx");
    res.send(buffer);

  } catch (err) {
    console.error("Error:", err);
    res.status(500).send("Server error: " + err.message);
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});