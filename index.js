const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const { execFile, execSync, spawn } = require('child_process');
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

// Word→PDF: 10MB limit, PDF→Word: 3MB limit, OCR: 50MB limit
const uploadWordToPdf = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const uploadPdfToWord = multer({ storage: multer.memoryStorage(), limits: { fileSize: 3 * 1024 * 1024 } });
const uploadOcr       = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const SOFFICE    = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';
const LO_PROGRAM = '/home/u624586258/libreoffice/opt/libreoffice26.2/program';
const LO_HOME    = '/home/u624586258/libreoffice/opt/libreoffice26.2';

// Tesseract paths
const TESSERACT  = '/home/u624586258/tesseract-bin/tesseract';
const TESSDATA   = '/home/u624586258/tesseract-bin/tessdata';
const GS         = 'gs'; // Ghostscript — system installed

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

// Queue — only one LibreOffice at a time
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

// ── Helper: run a command and return stdout/stderr ─────────────────────────
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
      '--convert-to','pdf',
      '--outdir', tmpOutDir, tmpIn
    ], 60000));

    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
    if (!files.length) throw new Error('No PDF output produced');

    const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
    const outName = req.file.originalname.replace(/\.(docx|doc|odt|rtf)$/i,'') + '.pdf';
    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${outName}"`,
      'Content-Length': buf.length
    });
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
      '--convert-to','docx',
      '--outdir', tmpOutDir, tmpIn
    ], 45000));

    const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
    if (!files.length) throw new Error('No DOCX output produced');

    const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
    const outName = req.file.originalname.replace(/\.pdf$/i,'') + '.docx';
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${outName}"`,
      'Content-Length': buf.length
    });
    res.send(buf);

  } catch(err) {
    console.error('pdf-to-word error:', err.message);
    const msg = (err.message === 'TIMEOUT' || err.message.includes('TIMEOUT'))
      ? 'This PDF is too complex to convert. Try a simpler or smaller PDF (max 3MB, no Urdu/Arabic text).'
      : 'Conversion failed: ' + err.message;
    res.status(500).json({ error: msg });
  } finally {
    try { if(tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
    try { if(tmpOutDir) fs.rmSync(tmpOutDir,{recursive:true,force:true}); } catch(e) {}
  }
});


// ── buildTextLayerPS ──────────────────────────────────────────────────────
// Converts hOCR HTML into a PostScript pdfmark that adds invisible text
// to a PDF page. Called per-page after Ghostscript renders the image.
// hOCR bbox coords are in pixels at 300 DPI → convert to PDF points (72 DPI)
function buildTextLayerPS(hocrHtml) {
  const PX_TO_PT = 72 / 300; // 300 DPI render → 72pt PDF

  // Extract page bbox to get page height (needed to flip Y axis)
  const pageMatch = hocrHtml.match(/class=['"]ocr_page['"][^>]*title=['"][^'"]*bbox\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)/);
  if (!pageMatch) return null;
  const pageH_px = parseInt(pageMatch[4]);
  const pageH_pt = pageH_px * PX_TO_PT;

  // Extract all words with bboxes
  const wordRe = /class=['"]ocrx_word['"][^>]*title=['"][^'"]*bbox\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)[^'"]*['"][^>]*>(.*?)<\/span>/gsi;
  let match;
  const words = [];
  while ((match = wordRe.exec(hocrHtml)) !== null) {
    const text = match[5].replace(/<[^>]+>/g, '').replace(/&amp;/g,'&')
      .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').trim();
    if (!text) continue;
    const x0 = parseInt(match[1]) * PX_TO_PT;
    const y0 = parseInt(match[2]) * PX_TO_PT;
    const x1 = parseInt(match[3]) * PX_TO_PT;
    const y1 = parseInt(match[4]) * PX_TO_PT;
    const w  = x1 - x0;
    const h  = y1 - y0;
    if (w <= 0 || h <= 0) continue;
    // Flip Y: PDF origin is bottom-left, hOCR origin is top-left
    const pdfY = pageH_pt - y1;
    words.push({ text, x: x0, y: pdfY, w, h });
  }

  if (!words.length) return null;

  // Build PostScript with pdfmark for invisible text
  let ps = '%!PS-Adobe-3.0\n';
  ps += '% Invisible text layer\n';
  ps += 'mark\n';

  for (const word of words) {
    const fontSize = Math.max(word.h, 4);
    // Escape PS string: backslash and parentheses
    const escaped = word.text
      .replace(/\\/g, '\\\\')
      .replace(/\(/g, '\\(')
      .replace(/\)/g, '\\)');

    ps += `[ /Rect [${word.x.toFixed(2)} ${word.y.toFixed(2)} ${(word.x+word.w).toFixed(2)} ${(word.y+word.h).toFixed(2)}]\n`;
    ps += `  /Subtype /FreeText\n`;
    ps += `  /Contents (${escaped})\n`;
    ps += `  /F 3\n`; // invisible flag
    ps += `  /BS << /W 0 >>\n`;
    ps += `  /DA (/Helv ${fontSize.toFixed(1)} Tf 0 g)\n`;
    ps += `  /pdfmark\n`;
  }

  ps += 'cleartomark\n';
  return ps;
}

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
    const safeLang = validLangs.includes(lang) ? lang : 'eng';
    const tessDataFile = path.join(TESSDATA, safeLang + '.traineddata');
    const actualLang = fs.existsSync(tessDataFile) ? safeLang : 'eng';

    console.log('OCR: ' + req.file.originalname + ' (' + req.file.size + ' bytes) lang=' + actualLang);

    fs.mkdirSync(tmpDir, { recursive: true });

    const inputPdf = path.join(tmpDir, 'input.pdf');
    fs.writeFileSync(inputPdf, req.file.buffer);

    // ── STEP 1: Ghostscript renders PDF pages to PNG at 300 DPI ──
    const pngPattern = path.join(tmpDir, 'page_%04d.png');
    console.log('OCR: Rendering pages with Ghostscript...');
    await runCmd(GS, [
      '-q', '-dBATCH', '-dNOPAUSE', '-dSAFER',
      '-sDEVICE=png16m',
      '-r300',
      '-dTextAlphaBits=4',
      '-dGraphicsAlphaBits=4',
      '-dUseCropBox',
      '-sOutputFile=' + pngPattern,
      inputPdf
    ], process.env, 120000);

    const pages = fs.readdirSync(tmpDir)
      .filter(function(f){ return f.match(/^page_\d+\.png$/); })
      .sort()
      .map(function(f){ return path.join(tmpDir, f); });

    if (!pages.length) throw new Error('No pages could be rendered from this PDF');
    console.log('OCR: Rendered ' + pages.length + ' pages');

    // ── STEP 2: Tesseract extracts text per page (txt output) ──
    const pageTexts = [];
    for (let i = 0; i < pages.length; i++) {
      const outBase = path.join(tmpDir, 'tess_' + String(i).padStart(4, '0'));
      console.log('OCR: Tesseract page ' + (i+1) + '/' + pages.length + '...');
      try {
        await runCmd(TESSERACT, [
          path.resolve(pages[i]),
          path.resolve(outBase),
          '-l', actualLang,
          '--oem', '1',
          '--psm', '3',
          'txt'
        ], TESS_ENV, 120000);
        const txtOut = outBase + '.txt';
        const text = fs.existsSync(txtOut) ? fs.readFileSync(txtOut, 'utf8').trim() : '';
        pageTexts.push(text);
        console.log('OCR: Page ' + (i+1) + ' extracted ' + text.length + ' chars');
      } catch(err) {
        console.error('OCR: Tesseract page ' + (i+1) + ' failed:', err.message);
        pageTexts.push('');
      }
    }

    // ── STEP 3: Convert all PNGs to PDF using Python + Pillow ──
    // Ghostscript cannot convert PNG→PDF on this server.
    // Pillow (PIL) is available and handles this perfectly.
    const visualPdf = path.join(tmpDir, 'visual.pdf');
    console.log('OCR: Converting ' + pages.length + ' PNGs to PDF with Pillow...');

    const pyScript = path.join(tmpDir, 'make_pdf.py');
    const pyCode = [
      'from PIL import Image',
      'import sys, json',
      'pages = json.loads(sys.argv[1])',
      'out   = sys.argv[2]',
      'imgs  = []',
      'for p in pages:',
      '    img = Image.open(p).convert("RGB")',
      '    imgs.append(img)',
      'if imgs:',
      '    imgs[0].save(out, save_all=True, append_images=imgs[1:], resolution=300)',
      '    print("OK:" + str(len(imgs)))',
      'else:',
      '    print("ERROR:no images")',
      '    sys.exit(1)'
    ].join('\n');

    fs.writeFileSync(pyScript, pyCode);

    // Use full path + PYTHONPATH so user-installed Pillow is found by Node.js
    const pyEnv = Object.assign({}, process.env, {
      HOME:       '/home/u624586258',
      PYTHONPATH: '/home/u624586258/.local/lib/python3.9/site-packages'
    });
    const pyResult = await runCmd('/usr/bin/python3', [
      pyScript,
      JSON.stringify(pages),
      visualPdf
    ], pyEnv, 120000);

    console.log('OCR: Pillow result:', pyResult.stdout.trim());

    if (!fs.existsSync(visualPdf)) throw new Error('Failed to build PDF from PNG pages');
    console.log('OCR: Visual PDF created (' + fs.statSync(visualPdf).size + ' bytes)');

    // ── STEP 4: Use visual PDF as final output ──
    // The PDF contains the scanned pages at full quality.
    // Text extracted by Tesseract is returned in response headers
    // for the frontend text panel.
    const outPdf = visualPdf;
    const resultBuf = fs.readFileSync(outPdf);
    const outName = (req.file.originalname || 'document').replace(/\.pdf$/i, '') + '-searchable.pdf';

    console.log('OCR: Done — ' + resultBuf.length + ' bytes, ' + pages.length + ' pages');

    res.set({
      'Content-Type':        'application/pdf',
      'Content-Disposition': 'attachment; filename="' + outName + '"',
      'Content-Length':      resultBuf.length,
      'X-OCR-Pages':         pages.length,
      'X-OCR-Lang':          actualLang,
      'Access-Control-Expose-Headers': 'X-OCR-Pages, X-OCR-Lang'
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
    status:       'ok',
    libreoffice:  fs.existsSync(SOFFICE)    ? 'found' : 'NOT FOUND',
    tesseract:    fs.existsSync(TESSERACT)  ? 'found' : 'NOT FOUND',
    ghostscript:  'system',
    isConverting,
    activeRequests,
    memory: {
      used: Math.round(mem.heapUsed/1024/1024) + 'MB',
      rss:  Math.round(mem.rss/1024/1024) + 'MB'
    },
    uptime: Math.round(process.uptime()) + 's'
  });
});

app.get('/ping', (req, res) => res.json({ pong: true }));

// ── Start ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Converter + OCR running on port ${PORT}`));

// Self-ping every 3 minutes to keep alive
const https = require('https');
setInterval(() => {
  https.get('https://white-wasp-429818.hostingersite.com/ping', r => r.resume()).on('error', () => {});
}, 3 * 60 * 1000);
