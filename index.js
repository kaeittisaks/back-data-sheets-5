const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const cors = require('cors'); // Import CORS module

const app = express();
app.use(cors()); // Enable CORS for all routes
const upload = multer({ dest: 'uploads/' });


app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const { path } = req.file;

    const result = await mammoth.extractRawText({ path });
    const text = result.value;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');
    const rows = text.split('\n');

    rows.forEach((row) => {
      const columns = row.split('\t');
      worksheet.addRow(columns);
    });

    const excelFilePath = 'data.xlsx';
    await workbook.xlsx.writeFile(excelFilePath);

    res.download(excelFilePath);
  } catch (error) {
    console.error(error);
    res.status(500).send('เกิดข้อผิดพลาดในการอัปโหลดไฟล์');
  }
});


app.listen(4000, () => {
  console.log('เซิร์ฟเวอร์เริ่มต้นที่พอร์ต 4000');
});


module.exports = app