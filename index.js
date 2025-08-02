const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const UPLOADS = path.join(__dirname, 'uploads');

// 🔹 Update this path if needed
const sofficePath = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe"`;

// Make sure uploads folder exists
if (!fs.existsSync(UPLOADS)) fs.mkdirSync(UPLOADS);

// Multer setup
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS),
  filename: (req, file, cb) => cb(null, `temp-${Date.now()}${path.extname(file.originalname)}`)
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xls') cb(null, true);
    else cb(new Error('Only .xls files are allowed'));
  }
});

// Route: Upload and Convert
app.post('/upload', upload.single('file'), (req, res) => {
  const inputFile = req.file.path;
  const baseFileName = path.basename(inputFile, path.extname(inputFile));
  const expectedConverted = path.join(UPLOADS, `${baseFileName}.xlsx`);
  const finalOutputFile = path.join(UPLOADS, `converted-${Date.now()}.xlsx`);

  const command = `${sofficePath} --headless --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir "${UPLOADS}" "${inputFile}"`;

  console.log(`Executing command: ${command}`);

  exec(command, (err, stdout, stderr) => {
    if (err || stderr) {
      console.error('❌ Conversion error:', stderr || err.message);
      return res.status(500).send('❌ Conversion failed.');
    }

    // Cleanup original
    if (fs.existsSync(inputFile)) fs.unlinkSync(inputFile);

    // Wait briefly to ensure file is written
    setTimeout(() => {
      if (!fs.existsSync(expectedConverted)) {
        return res.status(500).send('❌ Converted file not found.');
      }

      // Rename to timestamped file
      fs.renameSync(expectedConverted, finalOutputFile);

      res.status(200).send(`✅ Converted to: ${finalOutputFile}`);
    }, 500); // wait for LibreOffice to finish writing file
  });
});

app.listen(3000, () => {
  console.log('🚀 Server running at http://localhost:3000');
});
