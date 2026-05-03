const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { execFile } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

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

// Python path — uses user-installed pip packages
const PYTHON = '/usr/bin/python3';
const SCRIPT  = path.join(__dirname, 'convert.py');

function runPython(args, inputBuffer, timeoutMs = 60000) {
  return new Promise((resolve, reject) => {
    // Write input to a temp file
    const tmpIn  = path.join(os.tmpdir(), `in_${Date.now()}_${Math.random().toString(36).slice(2)}`);
    const tmpOut = path.join(os.tmpdir(), `out_${Date.now()}_${Math.random().toString(36).slice(2)}`);

    fs.writeFileSync(tmpIn, inputBuffer);

    const allArgs = [SCRIPT, ...args, tmpIn, tmpOut];

    const proc = execFile(PYTHON, allArgs, { timeout: timeoutMs }, (err, stdout, stderr) => {
      // Clean up input temp file
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
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (!['docx', 'doc'].includes(ext)) {
      return res.status(400).json({ error: 'Only .docx and .doc files are supported' });
    }

    const pageSize   = req.body.pageSize   || 'A4';
    const orientation= req.body.orientation|| 'portrait';
    const marginMm   = req.body.margin     || '20';

    console.log(`Converting Word→PDF: ${req.file.originalname} (${req.file.size} bytes)`);

    const pdfBuffer = await runPython(
      ['word-to-pdf', pageSize, orientation, marginMm],
      req.file.buffer
    );

    const outName = req.file.originalname.replace(/\.(docx|doc)$/i, '') + '.pdf';
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${outName}"`,
      'Content-Length': pdfBuffer.length
    });
    res.send(pdfBuffer);

  } catch (err) {
    console.error('word-to-pdf error:', err.message);
    res.status(500).json({ error: 'Conversion failed: ' + err.message });
  }
});

// ── POST /pdf-to-word ──────────────────────────────────────────────────────
app.post('/pdf-to-word', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (ext !== 'pdf') {
      return res.status(400).json({ error: 'Only PDF files are supported' });
    }

    console.log(`Converting PDF→Word: ${req.file.originalname} (${req.file.size} bytes)`);

    const docxBuffer = await runPython(
      ['pdf-to-word'],
      req.file.buffer
    );

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
  }
});

// ── Health check ───────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
  res.json({
    status: 'ok',
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
