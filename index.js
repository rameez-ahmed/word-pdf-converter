const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile, execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

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
const PYTHON3    = '/usr/bin/python3';
const PYTHONPATH = '/home/u624586258/.local/lib/python3.9/site-packages';

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

const PY_ENV = Object.assign({}, process.env, {
  HOME:       '/home/u624586258',
  PYTHONPATH: PYTHONPATH
});

// Create Tesseract config files on startup
function setupTesseractConfigs() {
  const configDir = path.join(TESSDATA, 'configs');
  try {
    fs.mkdirSync(configDir, { recursive: true });
    fs.writeFileSync(path.join(configDir, 'tsv'),  'tessedit_create_tsv 1\n');
    fs.writeFileSync(path.join(configDir, 'hocr'), 'tessedit_create_hocr 1\n');
    console.log('Tesseract configs ready');
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
app.post('/ocr-pdf', uploadOcr.single('pdf'), async (req, res) => {
  const tmpDir = path.join(os.tmpdir(), 'ocr_' + Date.now() + '_' + Math.random().toString(36).slice(2));
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

    console.log('OCR: ' + req.file.originalname + ' (' + req.file.size + ' bytes) lang=' + actualLang);
    fs.mkdirSync(tmpDir, { recursive: true });
    fs.writeFileSync(path.join(tmpDir, 'input.pdf'), req.file.buffer);

    // STEP 1: Ghostscript → PNG at 300 DPI
    console.log('OCR: Rendering pages...');
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

    if (!pages.length) throw new Error('No pages could be rendered');
    console.log('OCR: Rendered ' + pages.length + ' pages');

    // STEP 2: Tesseract → TSV per page (word bounding boxes)
    const pageTsvs  = [];
    const pageTexts = [];
    for (let i = 0; i < pages.length; i++) {
      const outBase = path.join(tmpDir, 'tess_' + String(i).padStart(4,'0'));
      console.log('OCR: Tesseract page ' + (i+1) + '/' + pages.length);
      try {
        await runCmd(TESSERACT, [
          path.resolve(pages[i]), path.resolve(outBase),
          '-l', actualLang, '--oem', '1', '--psm', '3', 'tsv'
        ], TESS_ENV, 120000);
        const tsv = outBase + '.tsv';
        if (fs.existsSync(tsv)) {
          pageTsvs.push(tsv);
          const words = fs.readFileSync(tsv, 'utf8').split('\n').slice(1)
            .map(l => l.split('\t'))
            .filter(p => p.length >= 12 && parseFloat(p[10]) > 0 && p[11] && p[11].trim())
            .map(p => p[11].trim());
          pageTexts.push(words.join(' '));
          console.log('OCR: Page ' + (i+1) + ' → ' + words.length + ' words');
        } else {
          pageTsvs.push(null);
          pageTexts.push('');
          console.log('OCR: Page ' + (i+1) + ' → no TSV');
        }
      } catch(e) {
        console.error('OCR: Tesseract p' + (i+1) + ' failed:', e.message);
        pageTsvs.push(null);
        pageTexts.push('');
      }
    }

    // STEP 3: Python builds searchable PDF
    // Key fix: use beginText() + setTextRenderMode(3) instead of _code injection
    // beginText() correctly handles graphics state after drawImage()
    console.log('OCR: Building searchable PDF...');

    const pyScript      = path.join(tmpDir, 'build_pdf.py');
    const pagesJsonFile = path.join(tmpDir, 'pages.json');
    const tsvsJsonFile  = path.join(tmpDir, 'tsvs.json');
    const finalPdf      = path.join(tmpDir, 'final.pdf');

    fs.writeFileSync(pagesJsonFile, JSON.stringify(pages));
    fs.writeFileSync(tsvsJsonFile,  JSON.stringify(pageTsvs));

    const py = [];
    py.push('import sys, json, os');
    py.push('from PIL import Image');
    py.push('from reportlab.pdfgen import canvas');
    py.push('from reportlab.lib.utils import ImageReader');
    py.push('');
    py.push('pages    = json.loads(open(sys.argv[1]).read())');
    py.push('tsvs     = json.loads(open(sys.argv[2]).read())');
    py.push('out_path = sys.argv[3]');
    py.push('SCALE    = 72.0 / 300.0');
    py.push('');
    py.push('def parse_tsv(tsv_file):');
    py.push('    words = []');
    py.push('    if not tsv_file or not os.path.exists(tsv_file): return words');
    py.push('    for line in open(tsv_file, encoding="utf-8").read().splitlines()[1:]:');
    py.push('        p = line.split("\\t")');
    py.push('        if len(p) < 12: continue');
    py.push('        try:');
    py.push('            if float(p[10]) < 0: continue');
    py.push('            text = p[11].strip()');
    py.push('            if not text: continue');
    py.push('            left, top, width, height = int(p[6]), int(p[7]), int(p[8]), int(p[9])');
    py.push('            if width > 0 and height > 0:');
    py.push('                words.append((text, left, top, width, height))');
    py.push('        except: pass');
    py.push('    return words');
    py.push('');
    py.push('img0   = Image.open(pages[0])');
    py.push('w0, h0 = img0.size');
    py.push('c = canvas.Canvas(out_path, pagesize=(w0*SCALE, h0*SCALE))');
    py.push('');
    py.push('for i, (png, tsv) in enumerate(zip(pages, tsvs)):');
    py.push('    img    = Image.open(png)');
    py.push('    iw, ih = img.size');
    py.push('    pw, ph = iw*SCALE, ih*SCALE');
    py.push('    c.setPageSize((pw, ph))');
    py.push('');
    py.push('    # Draw scanned image as background');
    py.push('    c.drawImage(ImageReader(img), 0, 0, width=pw, height=ph)');
    py.push('');
    py.push('    # Invisible text layer using beginText() + setTextRenderMode(3)');
    py.push('    # This is the correct ReportLab API that survives drawImage graphics state');
    py.push('    words = parse_tsv(tsv)');
    py.push('    if words:');
    py.push('        t = c.beginText()');
    py.push('        t.setTextRenderMode(3)  # 3 = invisible: no fill, no stroke');
    py.push('        for (text, left, top, width, height) in words:');
    py.push('            x  = left * SCALE');
    py.push('            y  = ph - (top + height) * SCALE');
    py.push('            fs = max(height * SCALE * 0.9, 1.0)');
    py.push('            t.setFont("Helvetica", fs)');
    py.push('            t.setTextOrigin(x, y)');
    py.push('            t.textLine(text)');
    py.push('        c.drawText(t)');
    py.push('    print("Page %d: %d words" % (i+1, len(words)))');
    py.push('    c.showPage()');
    py.push('');
    py.push('c.save()');
    py.push('print("PDF saved: " + out_path)');

    fs.writeFileSync(pyScript, py.join('\n'));

    const pyResult = await runCmd(PYTHON3, [
      pyScript, pagesJsonFile, tsvsJsonFile, finalPdf
    ], PY_ENV, 300000);

    console.log('OCR: Python:', pyResult.stdout.trim().split('\n').pop());
    if (!fs.existsSync(finalPdf)) throw new Error('Python failed to create PDF');
    console.log('OCR: Final PDF: ' + fs.statSync(finalPdf).size + ' bytes');

    const resultBuf = fs.readFileSync(finalPdf);
    const outName   = (req.file.originalname || 'document').replace(/\.pdf$/i,'') + '-searchable.pdf';
    const allText   = pageTexts.join('\n\n--- Page Break ---\n\n').trim();
    const textB64   = Buffer.from(allText.substring(0, 6000)).toString('base64');

    console.log('OCR: Done — ' + resultBuf.length + ' bytes, ' + pages.length + ' pages');

    res.set({
      'Content-Type':        'application/pdf',
      'Content-Disposition': 'attachment; filename="' + outName + '"',
      'Content-Length':      resultBuf.length,
      'X-OCR-Pages':         pages.length,
      'X-OCR-Lang':          actualLang,
      'X-OCR-Text':          textB64,
      'Access-Control-Expose-Headers': 'X-OCR-Pages, X-OCR-Lang, X-OCR-Text'
    });
    res.send(resultBuf);

  } catch(err) {
    console.error('ocr-pdf error:', err.message);
    res.status(500).json({ error: err.message || 'OCR processing failed' });
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch(e) {}
  }
});

// ── Status / ping ──────────────────────────────────────────────────────────
app.get('/status', (req, res) => {
  const mem = process.memoryUsage();
  res.json({
    status: 'ok',
    libreoffice: fs.existsSync(SOFFICE)   ? 'found' : 'NOT FOUND',
    tesseract:   fs.existsSync(TESSERACT) ? 'found' : 'NOT FOUND',
    ghostscript: 'system',
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
}, 3 * 60 * 1000);
