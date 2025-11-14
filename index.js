import express from 'express';
import multer from 'multer';
import cors from 'cors';
import { exec } from 'child_process';
import fs from 'fs';
import path from 'path';

const app = express();

// --- Middlewares --- //
app.use(cors());
app.use(express.json());

// Configure Multer for secure file uploads
const upload = multer({ dest: 'uploads/' });

// --- Helper Functions --- //
const cleanupFile = (filePath) => {
  if (fs.existsSync(filePath)) {
    fs.unlink(filePath, (err) => {
      if (err) console.error(`⚠️ Failed to clean up ${filePath}:`, err);
    });
  }
};

// --- Routes --- //
// 💡 Make sure LibreOffice is installed locally and added to system PATH so the headless conversion command works correctly.
app.get('/', (req, res) => {
  res.send('🚀 PPTX to PDF Converter API is live and ready to convert your slides!');
});

app.post('/convert', upload.single('file'), (req, res) => {
  const file = req.file;
  if (!file) return res.status(400).json({ error: 'No file uploaded.' });

  const inputPath = path.resolve(file.path);
  const outputDir = path.resolve('converted');
  const outputFileName = file.originalname.replace(/\.pptx$/i, '.pdf');
  const outputPath = path.join(outputDir, outputFileName);

  console.log(`📥 Received file: ${file.originalname}`);
  console.log('🧠 Starting LibreOffice conversion...');

  // Run LibreOffice in headless mode to convert PowerPoint to PDF
  exec(`libreoffice --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`, (error, stdout, stderr) => {
    if (error) {
      console.error('❌ Conversion error:', stderr || error);
      cleanupFile(inputPath);
      return res.status(500).json({ error: 'Conversion failed. Please ensure LibreOffice is installed and accessible in PATH.' });
    }

    if (fs.existsSync(outputPath)) {
      console.log('✅ Conversion successful:', outputPath);

      res.download(outputPath, outputFileName, (err) => {
        if (err) {
          console.error('⚠️ Download error:', err);
        } else {
          console.log(`🧹 Cleaning up temporary files for: ${file.originalname}`);
          cleanupFile(inputPath);
          cleanupFile(outputPath);
        }
      });
    } else {
      console.error('⚠️ Output file not found after conversion.');
      cleanupFile(inputPath);
      res.status(500).json({ error: 'Converted file not found.' });
    }
  });
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log('💡 Make sure LibreOffice is installed and accessible in your PATH');
});

