const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile } = require('child_process');
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

const SOFFICE = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';

// Queue — only one LibreOffice process at a time to prevent crashes
let loQueue = Promise.resolve();
function queueLibreOffice(fn) {
  const result = loQueue.then(fn, fn);
  loQueue = result.catch(() => {});
  return result;
}

function runLibreOffice(args, timeoutMs = 180000) {
  return new Promise((resolve, reject) => {
    const env = Object.assign({}, process.env, {
      HOME: '/home/u624586258',
      PATH: '/home/u624586258/libreoffice/opt/libreoffice26.2/program:' + (process.env.PATH || '')
    });
    execFile(SOFFICE, args, { timeout: timeoutMs, env }, (err, stdout, stderr) => {
      if (err) {
        console.error('LibreOffice error:', stderr || err.message);
        return reject(new Error(stderr || err.message));
      }
      resolve(stdout);
    });
  });
}

// ── POST /word-to-pdf ──────────────────────────────────────────────────────
app.post('/word-to-pdf', upload.single('file'), async (req, res) => {
  let tmpIn = null, tmpOutDir = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (!['docx', 'doc', 'odt', 'rtf'].includes(ext)) {
      return res.status(400).json({ error: 'Only .docx, .doc, .odt and .rtf files are supported' });
    }

    console.log(`Converting Word→PDF: ${req.file.originalname} (${req.file.size} bytes)`);

    const uniqueId = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn     = path.join(os.tmpdir(), uniqueId + '.' + ext);
    tmpOutDir = path.join(os.tmpdir(), uniqueId + '_out');

    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });

    await queueLibreOffice(() => runLibreOffice([
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--convert-to', 'pdf',
      '--outdir', tmpOutDir,
      tmpIn
    ]));

    // Find the output PDF
    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
    if (files.length === 0) throw new Error('LibreOffice did not produce a PDF output');
    const pdfPath = path.join(tmpOutDir, files[0]);
    const pdfBuffer = fs.readFileSync(pdfPath);
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
    try { if (tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
    try { if (tmpOutDir) fs.rmSync(tmpOutDir, { recursive: true, force: true }); } catch(e) {}
  }
});

// ── POST /pdf-to-word ──────────────────────────────────────────────────────
app.post('/pdf-to-word', upload.single('file'), async (req, res) => {
  let tmpIn = null, tmpOutDir = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (ext !== 'pdf') {
      return res.status(400).json({ error: 'Only PDF files are supported' });
    }

    console.log(`Converting PDF→Word: ${req.file.originalname} (${req.file.size} bytes)`);

    const uniqueId = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn     = path.join(os.tmpdir(), uniqueId + '.pdf');
    tmpOutDir = path.join(os.tmpdir(), uniqueId + '_out');

    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });

    await queueLibreOffice(() => runLibreOffice([
      '--headless',
      '--norestore',
      '--nofirststartwizard',
      '--infilter=writer_pdf_import',
      '--convert-to', 'docx',
      '--outdir', tmpOutDir,
      tmpIn
    ]));

    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
    if (files.length === 0) throw new Error('LibreOffice did not produce a DOCX output');
    const docxPath = path.join(tmpOutDir, files[0]);
    const docxBuffer = fs.readFileSync(docxPath);
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
    try { if (tmpOutDir) fs.rmSync(tmpOutDir, { recursive: true, force: true }); } catch(e) {}
  }
});

// ── Health check ───────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
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

// ── Self-ping every 4 minutes to keep server warm ─────────────────────────
const https = require('https');
setInterval(function(){
  try {
    https.get('https://white-wasp-429818.hostingersite.com/status', function(res){
      res.resume();
    }).on('error', function(){});
  } catch(e) {}
}, 4 * 60 * 1000);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Word/PDF converter running on port ${PORT}`));
