import express from 'express';
import multer from 'multer';
import cors from 'cors';
import { exec, execSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { PDFDocument } from 'pdf-lib';

const app = express();

// --- Middlewares --- //
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

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

// Detect LibreOffice command path
const getLibreOfficeCommand = () => {
  const platform = process.platform;
  
  // macOS paths
  if (platform === 'darwin') {
    const macPaths = [
      '/Applications/LibreOffice.app/Contents/MacOS/soffice',
      '/opt/homebrew/bin/soffice',
      '/usr/local/bin/soffice'
    ];
    
    for (const macPath of macPaths) {
      if (fs.existsSync(macPath)) {
        return macPath;
      }
    }
  }
  
  // Linux/Unix - try common paths
  const unixPaths = [
    '/usr/bin/libreoffice',
    '/usr/local/bin/libreoffice',
    '/opt/libreoffice*/program/soffice'
  ];
  
  // Try to find libreoffice/soffice in PATH
  try {
    const whichResult = execSync('which soffice 2>/dev/null || which libreoffice 2>/dev/null', { encoding: 'utf8' }).trim();
    if (whichResult) return whichResult;
  } catch (e) {
    // Not in PATH
  }
  
  // Default fallback
  return 'libreoffice';
};

const libreOfficeCmd = getLibreOfficeCommand();

// --- PDF Processing Functions --- //

/**
 * Get grid layout configuration for slides per page
 */
function getGridLayout(slidesPerPage) {
  const layouts = {
    2: { cols: 1, rows: 2 },
    3: { cols: 1, rows: 3 },
    4: { cols: 2, rows: 2 },
    6: { cols: 2, rows: 3 },
    9: { cols: 3, rows: 3 },
  };
  return layouts[slidesPerPage] || { cols: 1, rows: 1 };
}

/**
 * Merge multiple slides per page with proper scaling
 * @param {string} pdfPath - Path to input PDF
 * @param {number} slidesPerPage - Number of slides per page (2, 3, 4, 6, 9)
 * @returns {Promise<Uint8Array>} - Optimized PDF bytes
 */
async function mergeSlidesPerPage(pdfPath, slidesPerPage) {
  const existingPdfBytes = fs.readFileSync(pdfPath);
  const pdfDoc = await PDFDocument.load(existingPdfBytes);
  const pages = pdfDoc.getPages();
  const totalPages = pages.length;
  
  // Create new PDF document
  const newPdfDoc = await PDFDocument.create();
  
  // Calculate grid layout
  const grid = getGridLayout(slidesPerPage);
  const cols = grid.cols;
  const rows = grid.rows;
  
  // Standard US Letter size (8.5 x 11 inches) in points
  const pageWidth = 612;  // 8.5 * 72
  const pageHeight = 792; // 11 * 72
  
  // Calculate slide dimensions with padding
  const padding = 10;
  const slideWidth = (pageWidth - padding * (cols + 1)) / cols;
  const slideHeight = (pageHeight - padding * (rows + 1)) / rows;
  
  // Process pages in batches
  for (let i = 0; i < totalPages; i += slidesPerPage) {
    const newPage = newPdfDoc.addPage([pageWidth, pageHeight]);
    
    // Collect page indices for this batch
    const pageIndices = [];
    for (let j = 0; j < slidesPerPage && (i + j) < totalPages; j++) {
      pageIndices.push(i + j);
    }
    
    // Copy pages from source document
    const copiedPages = await newPdfDoc.copyPages(pdfDoc, pageIndices);
    
    // Draw each copied page onto the merged page
    for (let j = 0; j < copiedPages.length; j++) {
      const copiedPage = copiedPages[j];
      
      // Calculate position
      const col = j % cols;
      const row = Math.floor(j / cols);
      const x = padding + col * (slideWidth + padding);
      const y = pageHeight - (padding + (row + 1) * (slideHeight + padding));
      
      // Get source page dimensions
      const sourceSize = copiedPage.getSize();
      const sourceWidth = sourceSize.width;
      const sourceHeight = sourceSize.height;
      
      // Calculate scale to fit within allocated space
      const scaleX = slideWidth / sourceWidth;
      const scaleY = slideHeight / sourceHeight;
      const scale = Math.min(scaleX, scaleY);
      
      // Center the scaled slide
      const scaledWidth = sourceWidth * scale;
      const scaledHeight = sourceHeight * scale;
      const offsetX = (slideWidth - scaledWidth) / 2;
      const offsetY = (slideHeight - scaledHeight) / 2;
      
      // Draw the copied page onto the merged page
      newPage.drawPage(copiedPage, {
        x: x + offsetX,
        y: y + offsetY,
        width: scaledWidth,
        height: scaledHeight,
      });
    }
  }
  
  return await newPdfDoc.save();
}

/**
 * Compress PDF using Ghostscript
 * @param {string} inputPath - Path to input PDF
 * @param {string} outputPath - Path to output PDF
 * @returns {Promise<void>}
 */
function compressPdf(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    // Ghostscript compression command
    // /ebook = 150dpi (good balance between quality and size)
    const gsCommand = `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${outputPath}" "${inputPath}"`;
    
    exec(gsCommand, (error, stdout, stderr) => {
      if (error) {
        console.warn('⚠️ Ghostscript compression failed, using uncompressed PDF:', error.message);
        // Fallback: copy original if compression fails
        fs.copyFileSync(inputPath, outputPath);
        resolve();
      } else {
        console.log('✅ PDF compressed successfully');
        resolve();
      }
    });
  });
}

// --- Routes --- //
// 💡 Make sure LibreOffice is installed locally and added to system PATH so the headless conversion command works correctly.
// Static files are served from public directory, so root route serves index.html

app.post('/convert', upload.single('file'), async (req, res) => {
  const file = req.file;
  if (!file) return res.status(400).json({ error: 'No file uploaded.' });

  // Get slides per page from request (default to 1)
  const slidesPerPage = parseInt(req.body.slidesPerPage || req.query.slidesPerPage || '1');
  const validSlidesPerPage = [1, 2, 3, 4, 6, 9];
  if (!validSlidesPerPage.includes(slidesPerPage)) {
    cleanupFile(file.path);
    return res.status(400).json({ error: 'Invalid slidesPerPage. Must be 1, 2, 3, 4, 6, or 9.' });
  }

  const inputPath = path.resolve(file.path);
  const outputDir = path.resolve('converted');
  const baseFileName = file.originalname.replace(/\.pptx$/i, '');
  const tempPdfPath = path.join(outputDir, `${baseFileName}_temp.pdf`);
  const optimizedPdfPath = path.join(outputDir, `${baseFileName}_optimized.pdf`);
  const finalPdfPath = path.join(outputDir, `${baseFileName}.pdf`);

  console.log(`📥 Received file: ${file.originalname}`);
  console.log(`📊 Slides per page: ${slidesPerPage}`);

  // Ensure output directory exists
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  try {
    // Step 1: Convert PPTX to PDF using LibreOffice
    console.log('🧠 Step 1: Converting PPTX to PDF...');
    console.log(`📁 Input file: ${inputPath}`);
    console.log(`📁 Output directory: ${outputDir}`);
    
    await new Promise((resolve, reject) => {
      exec(`"${libreOfficeCmd}" --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`, (error, stdout, stderr) => {
        if (error) {
          console.error('LibreOffice stderr:', stderr);
          reject(new Error(`LibreOffice conversion failed: ${stderr || error.message}`));
        } else {
          console.log('LibreOffice stdout:', stdout);
          resolve();
        }
      });
    });

    // Find the converted PDF - LibreOffice uses the input filename but may sanitize it
    // Try multiple possible names
    const possibleNames = [
      `${baseFileName}.pdf`,
      path.basename(inputPath, path.extname(inputPath)) + '.pdf',
      file.originalname.replace(/\.pptx$/i, '.pdf')
    ];
    
    let libreOfficeOutput = null;
    for (const name of possibleNames) {
      const testPath = path.join(outputDir, name);
      if (fs.existsSync(testPath)) {
        libreOfficeOutput = testPath;
        console.log(`✅ Found converted PDF: ${name}`);
        break;
      }
    }
    
    // If still not found, search for any PDF file in the output directory
    if (!libreOfficeOutput) {
      console.log('🔍 Searching for PDF files in output directory...');
      const files = fs.readdirSync(outputDir);
      const pdfFiles = files.filter(f => f.toLowerCase().endsWith('.pdf'));
      console.log(`📄 Found PDF files: ${pdfFiles.join(', ')}`);
      
      if (pdfFiles.length > 0) {
        // Get the most recently modified PDF file
        const pdfPaths = pdfFiles.map(f => path.join(outputDir, f));
        pdfPaths.sort((a, b) => {
          const statA = fs.statSync(a);
          const statB = fs.statSync(b);
          return statB.mtime - statA.mtime;
        });
        libreOfficeOutput = pdfPaths[0];
        console.log(`✅ Using most recent PDF: ${path.basename(libreOfficeOutput)}`);
      }
    }
    
    if (!libreOfficeOutput || !fs.existsSync(libreOfficeOutput)) {
      throw new Error(`Converted PDF not found. Searched for: ${possibleNames.join(', ')}`);
    }

    console.log('✅ LibreOffice conversion successful');

    // Step 2: Merge slides per page (if slidesPerPage > 1)
    let pdfToCompress = libreOfficeOutput;
    if (slidesPerPage > 1) {
      console.log(`📄 Step 2: Merging ${slidesPerPage} slides per page...`);
      const mergedPdfBytes = await mergeSlidesPerPage(libreOfficeOutput, slidesPerPage);
      fs.writeFileSync(tempPdfPath, mergedPdfBytes);
      pdfToCompress = tempPdfPath;
      console.log('✅ Slides merged successfully');
    }

    // Step 3: Compress PDF
    console.log('🗜️ Step 3: Compressing PDF...');
    await compressPdf(pdfToCompress, finalPdfPath);

    // Cleanup temporary files
    cleanupFile(inputPath);
    if (fs.existsSync(libreOfficeOutput) && libreOfficeOutput !== finalPdfPath) {
      cleanupFile(libreOfficeOutput);
    }
    if (fs.existsSync(tempPdfPath)) {
      cleanupFile(tempPdfPath);
    }
    if (fs.existsSync(optimizedPdfPath) && optimizedPdfPath !== finalPdfPath) {
      cleanupFile(optimizedPdfPath);
    }

    console.log('✅ Conversion complete:', finalPdfPath);

    // Send optimized PDF
    res.download(finalPdfPath, `${baseFileName}.pdf`, (err) => {
      if (err) {
        console.error('⚠️ Download error:', err);
      } else {
        // Cleanup after download
        setTimeout(() => cleanupFile(finalPdfPath), 1000);
      }
    });

  } catch (error) {
    console.error('❌ Conversion error:', error);
    cleanupFile(inputPath);
    if (fs.existsSync(tempPdfPath)) cleanupFile(tempPdfPath);
    if (fs.existsSync(optimizedPdfPath)) cleanupFile(optimizedPdfPath);
    if (fs.existsSync(finalPdfPath)) cleanupFile(finalPdfPath);
    res.status(500).json({ error: error.message || 'Conversion failed.' });
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log('💡 Make sure LibreOffice is installed and accessible in your PATH');
});

