const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile, execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');
const crypto = require('crypto');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

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

const uploadWordToPdf = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const uploadPdfToWord = multer({ storage: multer.memoryStorage(), limits: { fileSize: 3 * 1024 * 1024 } });
const uploadOcr       = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const SOFFICE    = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';
const LO_PROGRAM = '/home/u624586258/libreoffice/opt/libreoffice26.2/program';
const LO_HOME    = '/home/u624586258/libreoffice/opt/libreoffice26.2';
const TESSERACT  = '/home/u624586258/tesseract-bin/tesseract';
const TESSDATA   = '/home/u624586258/tesseract-bin/tessdata';
const GS         = 'gs';

const LO_ENV = Object.assign({}, process.env, {
  HOME:            '/home/u624586258',
  PATH:            LO_PROGRAM + ':' + (process.env.PATH || ''),
  PYTHONHOME:      '',
  PYTHONPATH:      '',
  LD_LIBRARY_PATH: LO_PROGRAM + ':' + LO_HOME + '/ure-link/lib'
});

const TESS_ENV = Object.assign({}, process.env, {
  TESSDATA_PREFIX: TESSDATA,
  HOME:            '/home/u624586258',
  PATH:            '/home/u624586258/tesseract-bin:' + (process.env.PATH || '')
});

// ── Job store (in-memory) ──
// Stores job state: pending → processing → done / error
// Jobs are cleaned up after 10 minutes
const jobs = new Map();

function createJob(id) {
  jobs.set(id, {
    id,
    status:    'pending',
    progress:  0,
    step:      'Queued...',
    pages:     0,
    lang:      'eng',
    createdAt: Date.now()
  });
  return jobs.get(id);
}

function updateJob(id, data) {
  const job = jobs.get(id);
  if (job) Object.assign(job, data);
}

// Clean up old jobs every 5 minutes
setInterval(() => {
  const now = Date.now();
  for (const [id, job] of jobs.entries()) {
    if (now - job.createdAt > 10 * 60 * 1000) {
      // Clean up result file if exists
      if (job.resultPath && fs.existsSync(job.resultPath)) {
        try { fs.unlinkSync(job.resultPath); } catch(e) {}
      }
      if (job.tmpDir && fs.existsSync(job.tmpDir)) {
        try { fs.rmSync(job.tmpDir, { recursive: true, force: true }); } catch(e) {}
      }
      jobs.delete(id);
      console.log('Cleaned up job:', id);
    }
  }
}, 5 * 60 * 1000);

// Create Tesseract config files on startup
function setupTesseractConfigs() {
  const configDir = path.join(TESSDATA, 'configs');
  try {
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'tsv'),  'tessedit_create_tsv 1\n');
    fs.writeFileSync(path.join(configDir, 'hocr'), 'tessedit_create_hocr 1\n');
    fs.writeFileSync(path.join(configDir, 'pdf'),  'tessedit_create_pdf 1\n');
    fs.writeFileSync(path.join(configDir, 'txt'),  'tessedit_create_txt 1\n');
    console.log('Tesseract configs ready (tsv, hocr, pdf, txt)');
  } catch(e) {
    console.error('Tesseract config error:', e.message);
  }
}
setupTesseractConfigs();

let loQueue = Promise.resolve();
let isConverting = false;

function queueLibreOffice(fn) {
  const result = loQueue.then(() => {
    isConverting = true;
    return fn().finally(() => { isConverting = false; });
  });
  loQueue = result.catch(() => {
    isConverting = false;
    loQueue = Promise.resolve();
  });
  return result;
}

function killLibreOffice() {
  try { execSync('pkill -9 -f soffice 2>/dev/null || true', { timeout: 3000 }); } catch(e) {}
  try { execSync('rm -rf /tmp/*_out 2>/dev/null || true', { timeout: 3000 }); } catch(e) {}
}

function runCmd(bin, args, env, timeoutMs) {
  return new Promise((resolve, reject) => {
    execFile(bin, args, { timeout: timeoutMs, env: env, maxBuffer: 100 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) return reject(new Error(stderr || err.message));
        resolve({ stdout, stderr });
      }
    );
  });
}

function runLibreOffice(args, timeoutMs) {
  return new Promise((resolve, reject) => {
    execFile(SOFFICE, args, { timeout: timeoutMs, env: LO_ENV }, (err, stdout, stderr) => {
      if (err) {
        killLibreOffice();
        if (err.killed || err.code === null) return reject(new Error('TIMEOUT'));
        if (stderr && stderr.includes('Could not find platform')) return reject(new Error('STUCK'));
        return reject(new Error(stderr || err.message));
      }
      resolve(stdout);
    });
  });
}

async function convertFile(args, timeoutMs) {
  try {
    return await runLibreOffice(args, timeoutMs);
  } catch(err) {
    if (err.message === 'STUCK' || err.message === 'TIMEOUT') {
      console.log('LibreOffice stuck/timeout — killing and retrying once...');
      killLibreOffice();
      await new Promise(r => setTimeout(r, 2000));
      return await runLibreOffice(args, timeoutMs);
    }
    throw err;
  }
}

// ── POST /word-to-pdf ──────────────────────────────────────────────────────
app.post('/word-to-pdf', uploadWordToPdf.single('file'), async (req, res) => {
  let tmpIn = null, tmpOutDir = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (!['docx','doc','odt','rtf'].includes(ext))
      return res.status(400).json({ error: 'Only .docx, .doc, .odt and .rtf supported' });
    console.log(`Converting Word→PDF: ${req.file.originalname} (${req.file.size} bytes)`);
    const uid = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn     = path.join(os.tmpdir(), uid + '.' + ext);
    tmpOutDir = path.join(os.tmpdir(), uid + '_out');
    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });
    await queueLibreOffice(() => convertFile([
      '--headless','--norestore','--nofirststartwizard',
      '--convert-to','pdf','--outdir', tmpOutDir, tmpIn
    ], 60000));
    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
    if (!files.length) throw new Error('No PDF output produced');
    const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
    const outName = req.file.originalname.replace(/\.(docx|doc|odt|rtf)$/i,'') + '.pdf';
    res.set({ 'Content-Type':'application/pdf', 'Content-Disposition':`attachment; filename="${outName}"`, 'Content-Length': buf.length });
    res.send(buf);
  } catch(err) {
    console.error('word-to-pdf error:', err.message);
    res.status(500).json({ error: 'Conversion failed: ' + err.message });
  } finally {
    try { if(tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
    try { if(tmpOutDir) fs.rmSync(tmpOutDir,{recursive:true,force:true}); } catch(e) {}
  }
});

// ── POST /pdf-to-word ──────────────────────────────────────────────────────
app.post('/pdf-to-word', uploadPdfToWord.single('file'), async (req, res) => {
  let tmpIn = null, tmpOutDir = null;
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const ext = req.file.originalname.split('.').pop().toLowerCase();
    if (ext !== 'pdf') return res.status(400).json({ error: 'Only PDF files supported' });
    console.log(`Converting PDF→Word: ${req.file.originalname} (${req.file.size} bytes)`);
    const uid = Date.now() + '_' + Math.random().toString(36).slice(2);
    tmpIn     = path.join(os.tmpdir(), uid + '.pdf');
    tmpOutDir = path.join(os.tmpdir(), uid + '_out');
    fs.writeFileSync(tmpIn, req.file.buffer);
    fs.mkdirSync(tmpOutDir, { recursive: true });
    await queueLibreOffice(() => convertFile([
      '--headless','--norestore','--nofirststartwizard',
      '--infilter=writer_pdf_import',
      '--convert-to','docx','--outdir', tmpOutDir, tmpIn
    ], 45000));
    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
    if (!files.length) throw new Error('No DOCX output produced');
    const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
    const outName = req.file.originalname.replace(/\.pdf$/i,'') + '.docx';
    res.set({ 'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition':`attachment; filename="${outName}"`, 'Content-Length': buf.length });
    res.send(buf);
  } catch(err) {
    console.error('pdf-to-word error:', err.message);
    const msg = (err.message === 'TIMEOUT' || err.message.includes('TIMEOUT'))
      ? 'This PDF is too complex to convert. Try a simpler or smaller PDF.'
      : 'Conversion failed: ' + err.message;
    res.status(500).json({ error: msg });
  } finally {
    try { if(tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
    try { if(tmpOutDir) fs.rmSync(tmpOutDir,{recursive:true,force:true}); } catch(e) {}
  }
});

// ── POST /ocr-pdf ──────────────────────────────────────────────────────────
// Step 1: Upload file → get job_id immediately (no timeout risk)
app.post('/ocr-pdf', uploadOcr.single('pdf'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No PDF file uploaded' });
    const header = req.file.buffer.slice(0, 5).toString('ascii');
    if (!header.startsWith('%PDF')) return res.status(400).json({ error: 'Only PDF files are accepted' });

    const lang = (req.body.lang || 'eng').replace(/[^a-z_]/g, '');
    const validLangs = ['eng','ara','urd','fra','deu','spa','por','ita','nld','rus',
      'chi_sim','chi_tra','jpn','kor','hin','ben','tur','pol','ukr','vie',
      'tha','heb','fas','swe','nor','dan','fin','ces','ron','hun','ind','afr','swa'];
    const safeLang   = validLangs.includes(lang) ? lang : 'eng';
    const tessData   = path.join(TESSDATA, safeLang + '.traineddata');
    const actualLang = fs.existsSync(tessData) ? safeLang : 'eng';

    // Create job immediately
    const jobId  = crypto.randomBytes(16).toString('hex');
    const tmpDir = path.join(os.tmpdir(), 'ocr_' + jobId);
    fs.mkdirSync(tmpDir, { recursive: true });

    const job = createJob(jobId);
    job.tmpDir      = tmpDir;
    job.lang        = actualLang;
    job.filename    = req.file.originalname;
    job.status      = 'processing';

    // Save uploaded file
    fs.writeFileSync(path.join(tmpDir, 'input.pdf'), req.file.buffer);

    // Return job ID immediately — browser won't timeout
    res.json({ jobId, status: 'processing' });

    // Process in background (no await — fire and forget)
    processOcrJob(jobId, tmpDir, actualLang, req.file.originalname).catch(err => {
      console.error('OCR job error:', err.message);
      updateJob(jobId, { status: 'error', error: err.message });
    });

  } catch(err) {
    console.error('ocr-pdf submit error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /ocr-status/:jobId ──────────────────────────────────────────────────
// Step 2: Browser polls this every 3 seconds to check progress
app.get('/ocr-status/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ error: 'Job not found' });

  if (job.status === 'done') {
    return res.json({
      status:   'done',
      progress: 100,
      step:     'Complete!',
      pages:    job.pages,
      lang:     job.lang,
      size:     job.size,
      text:     job.text || ''
    });
  }

  if (job.status === 'error') {
    return res.json({ status: 'error', error: job.error });
  }

  res.json({
    status:   job.status,
    progress: job.progress,
    step:     job.step,
    pages:    job.pages
  });
});

// ── GET /ocr-download/:jobId ───────────────────────────────────────────────
// Step 3: Browser downloads the result PDF
app.get('/ocr-download/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ error: 'Job not found' });
  if (job.status !== 'done') return res.status(400).json({ error: 'Job not complete' });
  if (!job.resultPath || !fs.existsSync(job.resultPath)) {
    return res.status(404).json({ error: 'Result file not found' });
  }

  const buf     = fs.readFileSync(job.resultPath);
  const outName = (job.filename || 'document').replace(/\.pdf$/i,'') + '-searchable.pdf';

  res.set({
    'Content-Type':        'application/pdf',
    'Content-Disposition': 'attachment; filename="' + outName + '"',
    'Content-Length':      buf.length
  });
  res.send(buf);
});

// ── Background OCR processor ───────────────────────────────────────────────
async function processOcrJob(jobId, tmpDir, actualLang, filename) {
  console.log('OCR job start:', jobId, filename, 'lang=' + actualLang);

  try {
    // STEP 1: Ghostscript → PNG pages
    updateJob(jobId, { progress: 5, step: 'Rendering pages at 300 DPI...' });
    await runCmd(GS, [
      '-q', '-dBATCH', '-dNOPAUSE', '-dSAFER',
      '-sDEVICE=png16m', '-r300',
      '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4', '-dUseCropBox',
      '-sOutputFile=' + path.join(tmpDir, 'page_%04d.png'),
      path.join(tmpDir, 'input.pdf')
    ], process.env, 120000);

    const pages = fs.readdirSync(tmpDir)
      .filter(f => f.match(/^page_\d+\.png$/))
      .sort()
      .map(f => path.join(tmpDir, f));

    if (!pages.length) throw new Error('No pages could be rendered from this PDF');
    if (pages.length > 50) throw new Error('This PDF has ' + pages.length + ' pages. Maximum is 50 pages. Please split the PDF first.');

    updateJob(jobId, { progress: 15, step: 'Rendered ' + pages.length + ' pages. Running OCR...', pages: pages.length });
    console.log('OCR job', jobId, '— rendered', pages.length, 'pages');

    // STEP 2: Tesseract → PDF per page
    const pagePdfs  = [];
    const pageTexts = [];

    for (let i = 0; i < pages.length; i++) {
      const outBase = path.join(tmpDir, 'tess_' + String(i).padStart(4,'0'));
      const pct     = 15 + Math.round(((i+1) / pages.length) * 65);
      updateJob(jobId, {
        progress: pct,
        step: 'OCR page ' + (i+1) + ' of ' + pages.length + '...'
      });

      try {
        await runCmd(TESSERACT, [
          path.resolve(pages[i]),
          path.resolve(outBase),
          '-l', actualLang,
          '--oem', '1', '--psm', '3',
          'pdf', 'txt'
        ], TESS_ENV, 120000);

        const pdfOut = outBase + '.pdf';
        const txtOut = outBase + '.txt';
        if (fs.existsSync(pdfOut)) pagePdfs.push(pdfOut);
        if (fs.existsSync(txtOut)) pageTexts.push(fs.readFileSync(txtOut, 'utf8').trim());
        else pageTexts.push('');

      } catch(e) {
        console.error('OCR job', jobId, 'page', (i+1), 'failed:', e.message);
        pageTexts.push('');
      }
    }

    if (!pagePdfs.length) throw new Error('OCR failed — no pages could be processed');

    // STEP 3: Ghostscript merge
    updateJob(jobId, { progress: 82, step: 'Merging ' + pagePdfs.length + ' pages...' });
    const finalPdf = path.join(tmpDir, 'final.pdf');
    await runCmd(GS, [
      '-q', '-dBATCH', '-dNOPAUSE', '-dSAFER',
      '-sDEVICE=pdfwrite',
      '-dCompatibilityLevel=1.5',
      '-dPDFSETTINGS=/default',
      '-sOutputFile=' + finalPdf,
      ...pagePdfs
    ], process.env, 180000);

    if (!fs.existsSync(finalPdf)) throw new Error('Failed to merge pages');

    const size    = fs.statSync(finalPdf).size;
    const allText = pageTexts.join('\n\n--- Page Break ---\n\n').trim();
    const textB64 = Buffer.from(allText.substring(0, 6000)).toString('base64');

    updateJob(jobId, {
      status:     'done',
      progress:   100,
      step:       'Complete!',
      resultPath: finalPdf,
      pages:      pages.length,
      size:       size,
      text:       textB64
    });

    console.log('OCR job done:', jobId, '—', size, 'bytes,', pages.length, 'pages');

  } catch(err) {
    console.error('OCR job failed:', jobId, err.message);
    updateJob(jobId, { status: 'error', error: err.message });
    // Clean up on error
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch(e) {}
  }
}

// ── Status / ping ──────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
  res.json({
    status:      'ok',
    libreoffice: fs.existsSync(SOFFICE)   ? 'found' : 'NOT FOUND',
    tesseract:   fs.existsSync(TESSERACT) ? 'found' : 'NOT FOUND',
    ghostscript: 'system',
    activeJobs:  [...jobs.values()].filter(j => j.status === 'processing').length,
    totalJobs:   jobs.size,
    isConverting, activeRequests,
    memory: { used: Math.round(mem.heapUsed/1024/1024)+'MB', rss: Math.round(mem.rss/1024/1024)+'MB' },
    uptime: Math.round(process.uptime()) + 's'
  });
});

app.get('/ping', (req, res) => res.json({ pong: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Converter + OCR running on port ' + PORT));

const https = require('https');
setInterval(() => {
  https.get('https://white-wasp-429818.hostingersite.com/ping', r => r.resume()).on('error',()=>{});
}, 60 * 1000);
