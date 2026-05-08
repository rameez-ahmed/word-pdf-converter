const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile, execSync } = require('child_process');
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
    console.log(`[${new Date().toISOString()}] Done in ${Date.now()-start}ms — Status: ${res.statusCode}`);
  });
  next();
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

const SOFFICE   = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';
const LO_HOME   = '/home/u624586258/libreoffice/opt/libreoffice26.2';
const LO_PROGRAM= '/home/u624586258/libreoffice/opt/libreoffice26.2/program';

// Environment with all required paths set
const LO_ENV = Object.assign({}, process.env, {
  HOME:           '/home/u624586258',
  PATH:           LO_PROGRAM + ':' + (process.env.PATH || ''),
  URE_BOOTSTRAP:  'vnd.sun.star.pathname:' + LO_PROGRAM + '/fundamentalrc',
  PYTHONHOME:     '',
  PYTHONPATH:     '',
  LD_LIBRARY_PATH: LO_PROGRAM + ':' + LO_HOME + '/ure-link/lib'
});

// Queue — only one LibreOffice at a time
let loQueue = Promise.resolve();
function queueLibreOffice(fn) {
  const result = loQueue.then(fn, fn);
  loQueue = result.catch(() => {});
  return result;
}

// Kill any stuck LibreOffice processes
function killStuckLibreOffice() {
  try {
    execSync('pkill -f "soffice" 2>/dev/null || true', { timeout: 5000 });
  } catch(e) {}
  // Wait a moment for processes to die
  return new Promise(r => setTimeout(r, 1000));
}

function runLibreOffice(args, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    execFile(SOFFICE, args, { timeout: timeoutMs, env: LO_ENV }, (err, stdout, stderr) => {
      if (err) {
        // Check if it's the "platform libraries" error — means LO got stuck
        if (stderr && stderr.includes('Could not find platform')) {
          return reject(new Error('STUCK'));
        }
        console.error('LibreOffice error:', stderr || err.message);
        return reject(new Error(stderr || err.message));
      }
      resolve(stdout);
    });
  });
}

async function convertWithLibreOffice(args, timeoutMs) {
  try {
    return await runLibreOffice(args, timeoutMs);
  } catch(err) {
    if (err.message === 'STUCK') {
      // Kill stuck processes and retry once
      console.log('LibreOffice stuck — killing and retrying...');
      await killStuckLibreOffice();
      return await runLibreOffice(args, timeoutMs);
    }
    throw err;
  }
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

    await queueLibreOffice(() => convertWithLibreOffice([
      '--headless', '--norestore', '--nofirststartwizard',
      '--convert-to', 'pdf',
      '--outdir', tmpOutDir,
      tmpIn
    ]));

    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
    if (files.length === 0) throw new Error('LibreOffice did not produce a PDF');

    const pdfBuffer = fs.readFileSync(path.join(tmpOutDir, files[0]));
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

    await queueLibreOffice(() => convertWithLibreOffice([
      '--headless', '--norestore', '--nofirststartwizard',
      '--infilter=writer_pdf_import',
      '--convert-to', 'docx',
      '--outdir', tmpOutDir,
      tmpIn
    ]));

    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
    if (files.length === 0) throw new Error('LibreOffice did not produce a DOCX');

    const docxBuffer = fs.readFileSync(path.join(tmpOutDir, files[0]));
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

// ── Health / ping ──────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
  res.json({
    status: 'ok',
    libreoffice: fs.existsSync(SOFFICE) ? 'found' : 'NOT FOUND',
    activeRequests,
    memory: {
      used:  Math.round(mem.heapUsed  / 1024 / 1024) + 'MB',
      total: Math.round(mem.heapTotal / 1024 / 1024) + 'MB',
      rss:   Math.round(mem.rss       / 1024 / 1024) + 'MB'
    },
    uptime: Math.round(process.uptime()) + ' seconds'
  });
});

app.get('/ping', (req, res) => res.json({ pong: true, uptime: Math.round(process.uptime()) }));

// ── Start ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Word/PDF converter running on port ${PORT}`));

// Self-ping every 3 minutes
const https = require('https');
setInterval(() => {
  https.get('https://white-wasp-429818.hostingersite.com/ping', res => res.resume())
       .on('error', () => {});
}, 3 * 60 * 1000);
