const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { spawn, exec } = require('child_process');
const os = require('os');

const app = express();
const port = 3000;

// Create necessary directories
const uploadsDir = path.join(__dirname, 'uploads');
const tempDir = path.join(__dirname, 'temp');

[uploadsDir, tempDir].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const sanitizedName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, '_');
    const timestamp = Date.now();
    cb(null, `${timestamp}_${sanitizedName}`);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/octet-stream'
    ];
    
    const isExcelFile = allowedTypes.includes(file.mimetype) || 
                       file.originalname.match(/\.(xls|xlsx)$/i);
    
    if (isExcelFile) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xls, .xlsx) are allowed!'), false);
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB limit
  }
});

// Function to find LibreOffice installation path
function findLibreOfficePath() {
  const possiblePaths = [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files\\LibreOffice\\program\\soffice.com',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com',
    '/usr/bin/libreoffice',
    '/usr/bin/soffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice'
  ];
  
  for (const path of possiblePaths) {
    if (fs.existsSync(path)) {
      return path;
    }
  }
  
  return null;
}

// Alternative conversion function using direct LibreOffice command
function convertWithLibreOffice(inputPath, outputDir) {
  return new Promise((resolve, reject) => {
    const libreOfficePath = findLibreOfficePath();
    
    if (!libreOfficePath) {
      return reject(new Error('LibreOffice not found. Please install LibreOffice first.'));
    }
    
    console.log('📍 Using LibreOffice at:', libreOfficePath);
    
    // LibreOffice command to convert file
    const args = [
      '--headless',
      '--convert-to', 'xlsx',
      '--outdir', outputDir,
      inputPath
    ];
    
    console.log('🔄 Running command:', libreOfficePath, args.join(' '));
    
    const process = spawn(libreOfficePath, args, {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true
    });
    
    let stdout = '';
    let stderr = '';
    
    process.stdout.on('data', (data) => {
      stdout += data.toString();
    });
    
    process.stderr.on('data', (data) => {
      stderr += data.toString();
    });
    
    process.on('close', (code) => {
      console.log('📤 LibreOffice exit code:', code);
      console.log('📋 Stdout:', stdout);
      console.log('❌ Stderr:', stderr);
      
      if (code === 0) {
        // Find the converted file
        const inputFileName = path.basename(inputPath, path.extname(inputPath));
        const outputPath = path.join(outputDir, `${inputFileName}.xlsx`);
        
        if (fs.existsSync(outputPath)) {
          resolve(outputPath);
        } else {
          // Sometimes LibreOffice creates files with different names
          const files = fs.readdirSync(outputDir).filter(f => f.endsWith('.xlsx'));
          if (files.length > 0) {
            resolve(path.join(outputDir, files[files.length - 1]));
          } else {
            reject(new Error('Converted file not found'));
          }
        }
      } else {
        reject(new Error(`LibreOffice conversion failed with code ${code}: ${stderr}`));
      }
    });
    
    process.on('error', (error) => {
      reject(new Error(`Failed to start LibreOffice: ${error.message}`));
    });
  });
}

// Fallback conversion using Node.js Excel libraries
async function convertWithNodeLibs(inputPath, outputPath) {
  try {
    const XLSX = require('xlsx');
    
    console.log('📚 Using XLSX library for conversion');
    
    // Read the input file
    const workbook = XLSX.readFile(inputPath);
    
    // Write as XLSX format
    XLSX.writeFile(workbook, outputPath);
    
    return outputPath;
  } catch (error) {
    throw new Error(`Node.js conversion failed: ${error.message}`);
  }
}

// Serve static files for download
app.use('/downloads', express.static(uploadsDir));

// Upload and convert endpoint
app.post('/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded.' });
  }

  const inputPath = req.file.path;
  const timestamp = Date.now();
  const outputFilename = `converted_${timestamp}.xlsx`;
  const outputPath = path.join(uploadsDir, outputFilename);

  try {
    // Verify file exists and is readable
    if (!fs.existsSync(inputPath)) {
      throw new Error('Uploaded file not found');
    }

    const stats = fs.statSync(inputPath);
    console.log('📥 Processing file:', req.file.filename);
    console.log('📏 File size:', stats.size, 'bytes');

    let convertedPath;
    
    try {
      // Method 1: Try LibreOffice direct command
      console.log('🔄 Attempting LibreOffice conversion...');
      convertedPath = await convertWithLibreOffice(inputPath, uploadsDir);
      
      // Rename to our desired filename
      if (convertedPath !== outputPath) {
        fs.renameSync(convertedPath, outputPath);
        convertedPath = outputPath;
      }
      
      console.log('✅ LibreOffice conversion successful!');
      
    } catch (libreOfficeError) {
      console.log('⚠️ LibreOffice conversion failed:', libreOfficeError.message);
      console.log('🔄 Trying Node.js library fallback...');
      
      try {
        // Method 2: Fallback to Node.js libraries
        convertedPath = await convertWithNodeLibs(inputPath, outputPath);
        console.log('✅ Node.js library conversion successful!');
        
      } catch (nodeError) {
        console.log('❌ Node.js conversion also failed:', nodeError.message);
        throw new Error(`All conversion methods failed. LibreOffice: ${libreOfficeError.message}, Node.js: ${nodeError.message}`);
      }
    }
    
    // Clean up input file
    fs.unlinkSync(inputPath);

    res.json({
      success: true,
      message: 'File converted successfully!',
      downloadUrl: `/downloads/${outputFilename}`,
      originalName: req.file.originalname,
      convertedName: outputFilename
    });

  } catch (error) {
    console.error('❌ Error during conversion:', error);
    
    // Clean up files on error
    try {
      if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
      if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
    } catch (cleanupError) {
      console.error('Error during cleanup:', cleanupError);
    }

    res.status(500).json({
      error: 'Conversion failed',
      details: error.message
    });
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  const libreOfficePath = findLibreOfficePath();
  
  res.json({
    status: 'ok',
    libreOfficeFound: !!libreOfficePath,
    libreOfficePath: libreOfficePath || 'Not found',
    platform: os.platform(),
    nodeVersion: process.version
  });
});

// Simple upload form for testing
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
        <title>XLS to XLSX Converter</title>
        <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
            .upload-area { border: 2px dashed #ccc; padding: 40px; text-align: center; margin: 20px 0; }
            .upload-area:hover { border-color: #999; }
            input[type="file"] { margin: 20px 0; }
            button { background: #007bff; color: white; padding: 10px 20px; border: none; cursor: pointer; margin: 5px; }
            button:hover { background: #0056b3; }
            .result { margin-top: 20px; padding: 15px; border-radius: 5px; }
            .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
            .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
            .info { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
            .health-check { margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 5px; }
        </style>
    </head>
    <body>
        <h1>📊 XLS to XLSX Converter (Enhanced)</h1>
        
        <div class="health-check">
            <h3>🔧 System Status</h3>
            <button onclick="checkHealth()">Check System Health</button>
            <div id="healthResult"></div>
        </div>
        
        <div class="upload-area">
            <form id="uploadForm" enctype="multipart/form-data">
                <h3>Select Excel file to convert:</h3>
                <input type="file" id="fileInput" name="file" accept=".xls,.xlsx" required>
                <br>
                <button type="submit">🔄 Convert to XLSX</button>
            </form>
        </div>
        
        <div id="result"></div>
        
        <div class="info">
            <h4>ℹ️ Conversion Methods:</h4>
            <p><strong>Primary:</strong> LibreOffice direct command (most reliable)</p>
            <p><strong>Fallback:</strong> Node.js XLSX library (if LibreOffice fails)</p>
            <p><strong>Note:</strong> If you get errors, make sure LibreOffice is installed!</p>
        </div>

        <script>
            async function checkHealth() {
                const healthResult = document.getElementById('healthResult');
                healthResult.innerHTML = 'Checking...';
                
                try {
                    const response = await fetch('/health');
                    const result = await response.json();
                    
                    healthResult.innerHTML = \`
                        <div style="margin-top: 10px;">
                            <p><strong>LibreOffice Found:</strong> \${result.libreOfficeFound ? '✅ Yes' : '❌ No'}</p>
                            <p><strong>Path:</strong> \${result.libreOfficePath}</p>
                            <p><strong>Platform:</strong> \${result.platform}</p>
                            <p><strong>Node Version:</strong> \${result.nodeVersion}</p>
                        </div>
                    \`;
                } catch (error) {
                    healthResult.innerHTML = \`<div style="color: red; margin-top: 10px;">Error: \${error.message}</div>\`;
                }
            }
            
            document.getElementById('uploadForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const formData = new FormData();
                const fileInput = document.getElementById('fileInput');
                const resultDiv = document.getElementById('result');
                
                if (!fileInput.files[0]) {
                    resultDiv.innerHTML = '<div class="error">Please select a file!</div>';
                    return;
                }
                
                formData.append('file', fileInput.files[0]);
                resultDiv.innerHTML = '<div class="info">🔄 Converting... This may take a moment...</div>';
                
                try {
                    const response = await fetch('/upload', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const result = await response.json();
                    
                    if (result.success) {
                        resultDiv.innerHTML = \`
                            <div class="success">
                                <h4>✅ Conversion Successful!</h4>
                                <p><strong>Original:</strong> \${result.originalName}</p>
                                <p><strong>Converted:</strong> \${result.convertedName}</p>
                                <p><a href="\${result.downloadUrl}" download>📥 Download XLSX File</a></p>
                            </div>
                        \`;
                    } else {
                        resultDiv.innerHTML = \`<div class="error">❌ Error: \${result.error}<br><small>\${result.details}</small></div>\`;
                    }
                } catch (error) {
                    resultDiv.innerHTML = \`<div class="error">❌ Network error: \${error.message}</div>\`;
                }
            });
            
            // Check health on page load
            window.onload = checkHealth;
        </script>
    </body>
    </html>
  `);
});

// Error handling middleware
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ error: 'File too large. Maximum size is 10MB.' });
    }
  }
  
  console.error('Unhandled error:', error);
  res.status(500).json({ error: error.message || 'Internal server error' });
});

app.listen(port, () => {
  console.log('🚀 Server running at http://localhost:' + port);
  console.log('📁 Upload directory:', uploadsDir);
  console.log('🔧 Temp directory:', tempDir);
  
  const libreOfficePath = findLibreOfficePath();
  if (libreOfficePath) {
    console.log('✅ LibreOffice found at:', libreOfficePath);
  } else {
    console.log('⚠️ LibreOffice not found. Please install LibreOffice for best results.');
    console.log('📥 Download from: https://www.libreoffice.org/download/download/');
  }
});