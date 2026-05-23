const express = require('express');
const multer = require('multer');
const cors = require('cors');
const { execFile, execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

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
        console.log(`[${new Date().toISOString()}] Done in ${Date.now() - start}ms — Status: ${res.statusCode}`);
    });
    next();
});

const uploadWordToPdf = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const uploadPdfToWord = multer({ storage: multer.memoryStorage(), limits: { fileSize: 3 * 1024 * 1024 } });
const uploadOcr = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ================== PATHS & ENV ==================
const SOFFICE = '/home/u624586258/libreoffice/opt/libreoffice26.2/program/soffice';
const LO_PROGRAM = '/home/u624586258/libreoffice/opt/libreoffice26.2/program';
const LO_HOME = '/home/u624586258/libreoffice/opt/libreoffice26.2';
const TESSERACT = '/home/u624586258/tesseract-bin/tesseract';
const TESSDATA = '/home/u624586258/tesseract-bin/tessdata';
const GS = 'gs';
const PYTHON3 = '/usr/bin/python3';
const PYTHONPATH = '/home/u624586258/.local/lib/python3.9/site-packages';

const LO_ENV = Object.assign({}, process.env, {
    HOME: '/home/u624586258',
    PATH: LO_PROGRAM + ':' + (process.env.PATH || ''),
    PYTHONHOME: '',
    PYTHONPATH: '',
    LD_LIBRARY_PATH: LO_PROGRAM + ':' + LO_HOME + '/ure-link/lib'
});

const TESS_ENV = Object.assign({}, process.env, {
    TESSDATA_PREFIX: TESSDATA,
    HOME: '/home/u624586258',
    PATH: '/home/u624586258/tesseract-bin:' + (process.env.PATH || '')
});

const PY_ENV = Object.assign({}, process.env, {
    HOME: '/home/u624586258',
    PYTHONPATH: PYTHONPATH
});

// Setup Tesseract
function setupTesseractConfigs() {
    const configDir = path.join(TESSDATA, 'configs');
    try {
        fs.mkdirSync(configDir, { recursive: true });
        fs.writeFileSync(path.join(configDir, 'tsv'), 'tessedit_create_tsv 1\n');
        console.log('Tesseract configs ready');
    } catch (e) {
        console.error('Tesseract config error:', e.message);
    }
}
setupTesseractConfigs();

// LibreOffice Queue
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
    try { execSync('pkill -9 -f soffice 2>/dev/null || true', { timeout: 3000 }); } catch (e) {}
    try { execSync('rm -rf /tmp/*_out 2>/dev/null || true', { timeout: 3000 }); } catch (e) {}
}

function runCmd(bin, args, env, timeoutMs) {
    return new Promise((resolve, reject) => {
        execFile(bin, args, { timeout: timeoutMs, env: env, maxBuffer: 100 * 1024 * 1024 }, (err, stdout, stderr) => {
            if (err) return reject(new Error(stderr || err.message));
            resolve({ stdout, stderr });
        });
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
    } catch (err) {
        if (err.message === 'STUCK' || err.message === 'TIMEOUT') {
            console.log('LibreOffice stuck/timeout — killing and retrying...');
            killLibreOffice();
            await new Promise(r => setTimeout(r, 2000));
            return await runLibreOffice(args, timeoutMs);
        }
        throw err;
    }
}

// ── WORD TO PDF (Unchanged) ─────────────────────────────────────
app.post('/word-to-pdf', uploadWordToPdf.single('file'), async (req, res) => {
    let tmpIn = null, tmpOutDir = null;
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        const ext = req.file.originalname.split('.').pop().toLowerCase();
        if (!['docx','doc','odt','rtf'].includes(ext)) return res.status(400).json({ error: 'Only .docx, .doc, .odt and .rtf supported' });

        const uid = Date.now() + '_' + Math.random().toString(36).slice(2);
        tmpIn = path.join(os.tmpdir(), uid + '.' + ext);
        tmpOutDir = path.join(os.tmpdir(), uid + '_out');

        fs.writeFileSync(tmpIn, req.file.buffer);
        fs.mkdirSync(tmpOutDir, { recursive: true });

        await queueLibreOffice(() => convertFile(['--headless','--norestore','--nofirststartwizard','--convert-to','pdf','--outdir', tmpOutDir, tmpIn], 60000));

        const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.pdf'));
        if (!files.length) throw new Error('No PDF output');

        const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
        const outName = req.file.originalname.replace(/\.(docx|doc|odt|rtf)$/i,'') + '.pdf';

        res.set({ 'Content-Type':'application/pdf', 'Content-Disposition': `attachment; filename="${outName}"`, 'Content-Length': buf.length });
        res.send(buf);
    } catch(err) {
        console.error('word-to-pdf error:', err.message);
        res.status(500).json({ error: 'Conversion failed' });
    } finally {
        try { if(tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
        try { if(tmpOutDir) fs.rmSync(tmpOutDir,{recursive:true,force:true}); } catch(e) {}
    }
});

// ── PDF TO WORD (Unchanged) ─────────────────────────────────────
app.post('/pdf-to-word', uploadPdfToWord.single('file'), async (req, res) => {
    let tmpIn = null, tmpOutDir = null;
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        if (req.file.originalname.split('.').pop().toLowerCase() !== 'pdf') return res.status(400).json({ error: 'Only PDF supported' });

        const uid = Date.now() + '_' + Math.random().toString(36).slice(2);
        tmpIn = path.join(os.tmpdir(), uid + '.pdf');
        tmpOutDir = path.join(os.tmpdir(), uid + '_out');

        fs.writeFileSync(tmpIn, req.file.buffer);
        fs.mkdirSync(tmpOutDir, { recursive: true });

        await queueLibreOffice(() => convertFile(['--headless','--norestore','--nofirststartwizard','--infilter=writer_pdf_import','--convert-to','docx','--outdir', tmpOutDir, tmpIn], 45000));

        const files = fs.readdirSync(tmpOutDir).filter(f => f.endsWith('.docx'));
        if (!files.length) throw new Error('No DOCX output');

        const buf = fs.readFileSync(path.join(tmpOutDir, files[0]));
        const outName = req.file.originalname.replace(/\.pdf$/i,'') + '.docx';

        res.set({ 'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Content-Disposition': `attachment; filename="${outName}"`, 'Content-Length': buf.length });
        res.send(buf);
    } catch(err) {
        console.error('pdf-to-word error:', err.message);
        res.status(500).json({ error: 'Conversion failed' });
    } finally {
        try { if(tmpIn) fs.unlinkSync(tmpIn); } catch(e) {}
        try { if(tmpOutDir) fs.rmSync(tmpOutDir,{recursive:true,force:true}); } catch(e) {}
    }
});

// ── OCR PDF (NEW - Using PyMuPDF) ─────────────────────────────────
app.post('/ocr-pdf', uploadOcr.single('pdf'), async (req, res) => {
    const tmpDir = path.join(os.tmpdir(), 'ocr_' + Date.now() + '_' + Math.random().toString(36).slice(2));
    try {
        if (!req.file) return res.status(400).json({ error: 'No PDF file uploaded' });
        if (!req.file.buffer.slice(0, 5).toString('ascii').startsWith('%PDF')) return res.status(400).json({ error: 'Only PDF files accepted' });

        const lang = (req.body.lang || 'eng').replace(/[^a-z_]/g, '');
        const validLangs = ['eng','ara','urd','fra','deu','spa','por','ita','nld','rus','chi_sim','chi_tra','jpn','kor','hin','ben','tur','pol','ukr','vie','tha','heb','fas','swe','nor','dan','fin','ces','ron','hun','ind','afr','swa'];
        const actualLang = validLangs.includes(lang) && fs.existsSync(path.join(TESSDATA, lang + '.traineddata')) ? lang : 'eng';

        console.log(`OCR Started: ${req.file.originalname} (${req.file.size} bytes) Lang: ${actualLang}`);

        fs.mkdirSync(tmpDir, { recursive: true });
        fs.writeFileSync(path.join(tmpDir, 'input.pdf'), req.file.buffer);

        // Render pages
        await runCmd(GS, ['-q','-dBATCH','-dNOPAUSE','-dSAFER','-sDEVICE=png16m','-r400','-dTextAlphaBits=4','-dGraphicsAlphaBits=4','-dUseCropBox',`-sOutputFile=${path.join(tmpDir, 'page_%04d.png')}`, path.join(tmpDir, 'input.pdf')], process.env, 150000);

        const pages = fs.readdirSync(tmpDir).filter(f => f.match(/^page_\d+\.png$/)).sort().map(f => path.join(tmpDir, f));

        // Tesseract
        const pageTsvs = [], pageTexts = [];
        for (let i = 0; i < pages.length; i++) {
            const outBase = path.join(tmpDir, `tess_${String(i).padStart(4,'0')}`);
            await runCmd(TESSERACT, [pages[i], outBase, '-l', actualLang, '--oem', '3', '--psm', '6', 'tsv'], TESS_ENV, 90000);
            
            const tsv = outBase + '.tsv';
            if (fs.existsSync(tsv)) {
                pageTsvs.push(tsv);
                const text = fs.readFileSync(tsv, 'utf8').split('\n').slice(1).map(l => l.split('\t')[11]).filter(Boolean).join(' ');
                pageTexts.push(text.trim());
            } else {
                pageTsvs.push(null);
                pageTexts.push('');
            }
        }

        // Build PDF with PyMuPDF
        const pyScript = path.join(tmpDir, 'build_pdf.py');
        const pagesJson = path.join(tmpDir, 'pages.json');
        const tsvsJson = path.join(tmpDir, 'tsvs.json');
        const finalPdf = path.join(tmpDir, 'final.pdf');

        fs.writeFileSync(pagesJson, JSON.stringify(pages));
        fs.writeFileSync(tsvsJson, JSON.stringify(pageTsvs));

        const pyCode = `
import sys, json, os
import fitz
from PIL import Image

pages = json.loads(open(sys.argv[1]).read())
tsvs = json.loads(open(sys.argv[2]).read())
out_path = sys.argv[3]

def parse_tsv(tsv_path):
    if not tsv_path or not os.path.exists(tsv_path): return []
    words = []
    for line in open(tsv_path, encoding='utf-8').read().splitlines()[1:]:
        parts = line.split('\t')
        if len(parts) < 12: continue
        try:
            if float(parts[10]) < 30: continue
            text = parts[11].strip()
            if not text: continue
            x0 = int(parts[6])
            y0 = int(parts[7])
            x1 = x0 + int(parts[8])
            y1 = y0 + int(parts[9])
            words.append((text, x0, y0, x1, y1))
        except: pass
    return words

doc = fitz.open()
for i, (png_path, tsv_path) in enumerate(zip(pages, tsvs)):
    w, h = Image.open(png_path).size
    page = doc.new_page(width=w, height=h)
    page.insert_image(page.rect, filename=png_path)
    
    words = parse_tsv(tsv_path)
    if words:
        for text, x0, y0, x1, y1 in words:
            rect = fitz.Rect(x0, y0, x1, y1)
            page.insert_textbox(rect, text, fontsize=(y1-y0)*0.82, fontname="helv", render_mode=3)
    print(f"Page {i+1}: {len(words)} words")
    
doc.save(out_path, deflate=True, clean=True)
doc.close()
print("Searchable PDF created")
`;

        fs.writeFileSync(pyScript, pyCode);
        await runCmd(PYTHON3, [pyScript, pagesJson, tsvsJson, finalPdf], PY_ENV, 180000);

        const resultBuf = fs.readFileSync(finalPdf);
        const outName = (req.file.originalname || 'document').replace(/\.pdf$/i, '') + '-searchable.pdf';

        res.set({
            'Content-Type': 'application/pdf',
            'Content-Disposition': `attachment; filename="${outName}"`,
            'Content-Length': resultBuf.length
        });
        res.send(resultBuf);

    } catch (err) {
        console.error('OCR Error:', err.message);
        res.status(500).json({ error: 'OCR processing failed. Try a clearer document.' });
    } finally {
        try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch(e) {}
    }
});

// Status
app.get('/status', (req, res) => res.json({ status: 'ok' }));
app.get('/ping', (req, res) => res.json({ pong: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('OCR Server running on port ' + PORT));
