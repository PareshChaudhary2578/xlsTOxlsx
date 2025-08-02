const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const ConvertAPI = require('convertapi')('CEph3V3dCuW2qAeSaYgAe9FeVW2zzPgZ'); // Replace with your secret

const app = express();
const port = 3000;

// Create uploads folder if not exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// File upload config
const upload = multer({ dest: 'uploads/' });

// Convert Function using ConvertAPI
async function convertXlsToXlsx(inputPath, outputPath) {
  try {
    const result = await ConvertAPI.convert('xlsx', {
      File: inputPath
    }, 'xls');

    await result.file.save(outputPath);
    console.log(`✅ File converted to: ${outputPath}`);
    return true;
  } catch (error) {
    console.error('❌ Conversion failed:', error.message);
    return false;
  }
}

// Upload Route
app.post('/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('❌ No file uploaded.');
  }

  const filePath = req.file.path;
  const originalName = req.file.originalname;
  const ext = path.extname(originalName).toLowerCase();

  if (ext !== '.xls') {
    fs.unlinkSync(filePath);
    return res.status(400).send('❌ Only .xls files are allowed.');
  }

  // ✅ Rename uploaded file with .xls extension
  const renamedPath = `${filePath}.xls`;
  fs.renameSync(filePath, renamedPath);

  const outputFilename = `converted-${Date.now()}.xlsx`;
  const outputFilePath = path.join(__dirname, 'uploads', outputFilename);

  const success = await convertXlsToXlsx(renamedPath, outputFilePath);
  fs.unlinkSync(renamedPath); // remove uploaded .xls

  if (success) {
    res.json({
      message: '✅ File converted!',
      downloadUrl: `/uploads/${outputFilename}`
    });
  } else {
    res.status(500).send('❌ Conversion failed.');
  }
});


// Serve converted files
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// Start server
app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});
