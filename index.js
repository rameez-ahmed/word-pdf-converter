const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile, exec } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

let activeRequests = 0;

app.use((req, res, next) => {
  activeRequests++;
  const start = Date.now();
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path} — Active: ${activeRequests}`);
  res.on('finish', () => {
    activeRequests--;
    const duration = Date.now() - start;
    console.log(`[${new Date().toISOString()}] Done in ${duration}ms — Status: ${res.statusCode}`);
  });
  next();
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

// LibreOffice path — installed without sudo in home folder
const SOFFICE = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';

// Python path for PDF-to-Word fallback
const PYTHON = '/usr/bin/python3';
const CONVERT_PY = path.join(__dirname, 'convert.py');

function runLibreOffice(args, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    // Set HOME so LibreOffice can write its user profile
    const env = Object.assign({}, process.env, {
      HOME: '/home/u624586258',
      PATH: '/home/u624586258/libreoffice/opt/libreoffice26.2/program:' + (process.env.PATH || '')
    });

    const proc = execFile(SOFFICE, args, { timeout: timeoutMs, env }, (err, stdout, stderr) => {
      if (err) {
        console.error('LibreOffice error:', stderr || err.message);
        return reject(new Error(stderr || err.message));
      }
      resolve(stdout);
    });
  });
}

function runPython(args, inputBuffer, timeoutMs = 60000) {
  return new Promise((resolve, reject) => {
    const tmpIn  = path.join(os.tmpdir(), 'in_'  + Date.now() + '_' + Math.random().toString(36).slice(2));
    const tmpOut = path.join(os.tmpdir(), 'out_' + Date.now() + '_' + Math.random().toString(36).slice(2));

    fs.writeFileSync(tmpIn, inputBuffer);

    const allArgs = [CONVERT_PY, ...args, tmpIn, tmpOut];
    const env = Object.assign({}, process.env, {
      HOME: '/home/u624586258'
    });

    execFile(PYTHON, allArgs, { timeout: timeoutMs, env }, (err, stdout, stderr) => {
      try { fs.unlinkSync(tmpIn); } catch(e) {}
      if (err) {
        try { fs.unlinkSync(tmpOut); } catch(e) {}
        console.error('Python error:', stderr || err.message);
        return reject(new Error(stderr || err.message));
      }
      try {
        const outBuf = fs.readFileSync(tmpOut);
        fs.unlinkSync(tmpOut);
        resolve(outBuf);
      } catch(e) {
        reject(new Error('Output file not created: ' + e.message));
      }
    });
  });
}

// ── POST /word-to-pdf ──────────────────────────────────────────────────────
app.post('/word-to-pdf', upload.single('file'), async (req, res) => {
  let tmpIn = null, tmpOut = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (!['docx', 'doc', 'odt', 'rtf'].includes(ext)) {
      return res.status(400).json({ error: 'Only .docx, .doc, .odt and .rtf files are supported' });
    }

    console.log(`Converting Word→PDF: ${req.file.originalname} (${req.file.size} bytes)`);

    // Write input to temp file
    const tmpDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn  = path.join(tmpDir, uniqueId + '.' + ext);
    const tmpOutDir = path.join(tmpDir, uniqueId + '_out');

    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });

    // Run LibreOffice headless conversion
    await runLibreOffice([
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--convert-to', 'pdf',
      '--outdir', tmpOutDir,
      tmpIn
    ]);

    // Find the output PDF
    const baseName = path.basename(tmpIn, '.' + ext);
    tmpOut = path.join(tmpOutDir, baseName + '.pdf');

    if (!fs.existsSync(tmpOut)) {
      // Try to find any PDF in the output dir
      const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
      if (files.length === 0) throw new Error('LibreOffice did not produce a PDF output');
      tmpOut = path.join(tmpOutDir, files[0]);
    }

    const pdfBuffer = fs.readFileSync(tmpOut);
    const outName = req.file.originalname.replace(/\.(docx|doc|odt|rtf)$/i, '') + '.pdf';

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${outName}"`,
      'Content-Length': pdfBuffer.length
    });
    res.send(pdfBuffer);

  } catch (err) {
    console.error('word-to-pdf error:', err.message);
    res.status(500).json({ error: 'Conversion failed: ' + err.message });
  } finally {
    // Clean up temp files
    try { if (tmpIn)  fs.unlinkSync(tmpIn);  } catch(e) {}
    try {
      if (tmpOut) {
        const outDir = path.dirname(tmpOut);
        fs.rmSync(outDir, { recursive: true, force: true });
      }
    } catch(e) {}
  }
});

// ── POST /pdf-to-word ──────────────────────────────────────────────────────
app.post('/pdf-to-word', upload.single('file'), async (req, res) => {
  let tmpIn = null, tmpOut = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (ext !== 'pdf') {
      return res.status(400).json({ error: 'Only PDF files are supported' });
    }

    console.log(`Converting PDF→Word: ${req.file.originalname} (${req.file.size} bytes)`);

    // Write input to temp file
    const tmpDir = os.tmpdir();
    const uniqueId = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn = path.join(tmpDir, uniqueId + '.pdf');
    const tmpOutDir = path.join(tmpDir, uniqueId + '_out');

    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });

    // LibreOffice: PDF → DOCX
    await runLibreOffice([
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--infilter=writer_pdf_import',
      '--convert-to', 'docx',
      '--outdir', tmpOutDir,
      tmpIn
    ]);

    // Find output docx
    const baseName = path.basename(tmpIn, '.pdf');
    tmpOut = path.join(tmpOutDir, baseName + '.docx');

    if (!fs.existsSync(tmpOut)) {
      const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
      if (files.length === 0) throw new Error('LibreOffice did not produce a DOCX output');
      tmpOut = path.join(tmpOutDir, files[0]);
    }

    const docxBuffer = fs.readFileSync(tmpOut);
    const outName = req.file.originalname.replace(/\.pdf$/i, '') + '.docx';

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${outName}"`,
      'Content-Length': docxBuffer.length
    });
    res.send(docxBuffer);

  } catch (err) {
    console.error('pdf-to-word error:', err.message);
    res.status(500).json({ error: 'Conversion failed: ' + err.message });
  } finally {
    try { if (tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
    try {
      if (tmpOut) {
        const outDir = path.dirname(tmpOut);
        fs.rmSync(outDir, { recursive: true, force: true });
      }
    } catch(e) {}
  }
});

// ── Health check ───────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
  // Check if LibreOffice exists
  const loExists = fs.existsSync(SOFFICE);
  res.json({
    status: 'ok',
    libreoffice: loExists ? 'found' : 'NOT FOUND',
    activeRequests,
    memory: {
      used:  Math.round(mem.heapUsed  / 1024 / 1024) + 'MB',
      total: Math.round(mem.heapTotal / 1024 / 1024) + 'MB',
      rss:   Math.round(mem.rss       / 1024 / 1024) + 'MB'
    },
    uptime: Math.round(process.uptime()) + ' seconds'
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Word/PDF converter running on port ${PORT}`));
