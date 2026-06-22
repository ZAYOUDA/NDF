require('dotenv').config();
const express    = require('express');
const multer     = require('multer');
const ExcelJS    = require('exceljs');
const PDFDocument = require('pdfkit');
const path       = require('path');
const fs         = require('fs');
const https      = require('https');

const app  = express();
const PORT = process.env.PORT || 3000;
const USE_SUPABASE = !!process.env.SUPABASE_URL;

console.log(`Mode : ${USE_SUPABASE ? '🟢 PROD (Supabase)' : '🟡 DEV (local JSON)'}`);

// ── Multer ────────────────────────────────────────────────
const imgFilter  = (req, file, cb) =>
  file.mimetype.startsWith('image/') ? cb(null, true) : cb(new Error('Images uniquement'));
const xlsxFilter = (req, file, cb) =>
  file.mimetype.includes('spreadsheet') || file.originalname.endsWith('.xlsx') ? cb(null, true) : cb(new Error('Fichier Excel uniquement'));

let upload;
if (USE_SUPABASE) {
  upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20*1024*1024 }, fileFilter: imgFilter });
} else {
  const UPLOAD_DIR = path.join(__dirname, 'uploads');
  if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);
  upload = multer({
    storage: multer.diskStorage({
      destination: UPLOAD_DIR,
      filename: (req, file, cb) => cb(null, `${crypto.randomUUID()}${path.extname(file.originalname)||'.jpg'}`),
    }),
    limits: { fileSize: 20*1024*1024 },
    fileFilter: imgFilter,
  });
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
if (!USE_SUPABASE) app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// ── Supabase (PROD) ───────────────────────────────────────
let supabase;
const BUCKET = 'receipts';
if (USE_SUPABASE) {
  const { createClient } = require('@supabase/supabase-js');
  supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);
}

// ── Local DB (DEV) ────────────────────────────────────────
const DATA_FILE = path.join(__dirname, 'data.json');
function readDB()    { try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); } catch { return { nextId: 1, expenses: [] }; } }
function writeDB(db) { fs.writeFileSync(DATA_FILE, JSON.stringify(db, null, 2)); }

// ── Helpers communs ───────────────────────────────────────
function computeTotals(vatLines, transport, repas) {
  const r = { tva_2_6: 0, tva_5_5: 0, tva_10: 0, tva_20: 0, total_ht: 0, total_ttc: 0 };
  r.total_ttc = (parseFloat(transport)||0) + (parseFloat(repas)||0);
  (vatLines||[]).forEach(line => {
    const ttc = parseFloat(line.ttc)||0, rate = parseFloat(line.rate)||0;
    const ht  = ttc / (1 + rate/100);
    r.total_ht += ht;
    if (rate === 2.6) r.tva_2_6 += ttc-ht;
    if (rate === 5.5) r.tva_5_5 += ttc-ht;
    if (rate === 10)  r.tva_10  += ttc-ht;
    if (rate === 20)  r.tva_20  += ttc-ht;
  });
  if (!vatLines||!vatLines.length) r.total_ht = r.total_ttc;
  return r;
}

function mapRow(row) {
  return {
    id:          row.id,
    description: row.description || '',
    client:      row.client      || '',
    projet:      row.projet      || '',
    mois:        row.mois        || '',
    categorie:   row.categorie   || '',
    transport:   parseFloat(row.transport)  || 0,
    repas:       parseFloat(row.repas)      || 0,
    commentaire: row.commentaire || '',
    vatLines:    row.vat_lines   || [],
    totalTTC:    parseFloat(row.total_ttc)  || 0,
    totalHT:     parseFloat(row.total_ht)   || 0,
    tva_2_6:     parseFloat(row.tva_2_6)   || 0,
    tva_5_5:     parseFloat(row.tva_5_5)   || 0,
    tva_10:      parseFloat(row.tva_10)    || 0,
    tva_20:      parseFloat(row.tva_20)    || 0,
    images:      row.images      || [],
    createdAt:   row.created_at  || row.createdAt,
  };
}

async function uploadImages(files) {
  if (USE_SUPABASE) {
    const urls = [];
    for (const file of files) {
      const ext = path.extname(file.originalname)||'.jpg';
      const filename = `${crypto.randomUUID()}${ext}`;
      const { error } = await supabase.storage.from(BUCKET).upload(filename, file.buffer, { contentType: file.mimetype, upsert: false });
      if (error) throw error;
      const { data } = supabase.storage.from(BUCKET).getPublicUrl(filename);
      urls.push(data.publicUrl);
    }
    return urls;
  } else {
    return (files||[]).map(f => `/uploads/${f.filename}`);
  }
}

async function deleteImages(urls) {
  if (USE_SUPABASE) {
    const paths = urls.map(u => u.split(`${BUCKET}/`)[1]).filter(Boolean);
    if (paths.length) await supabase.storage.from(BUCKET).remove(paths);
  } else {
    urls.forEach(url => { try { fs.unlinkSync(path.join(__dirname, url)); } catch {} });
  }
}

// ── API : liste ───────────────────────────────────────────
app.get('/api/expenses', async (req, res) => {
  if (USE_SUPABASE) {
    const { data, error } = await supabase.from('expenses').select('*').order('id', { ascending: true });
    if (error) return res.status(500).json({ error: error.message });
    return res.json(data.map(mapRow));
  }
  const db = readDB();
  res.json(db.expenses.map(mapRow));
});

// ── API : ajout ───────────────────────────────────────────
app.post('/api/expenses', upload.array('images', 20), async (req, res) => {
  const { description, client, projet, mois, categorie, transport, repas, commentaire, vatLines } = req.body;
  const parsedVat = JSON.parse(vatLines||'[]');
  const totals    = computeTotals(parsedVat, transport, repas);
  const images    = await uploadImages(req.files||[]);

  if (USE_SUPABASE) {
    const { data, error } = await supabase.from('expenses').insert([{
      description: description||'', client: client||'', projet: projet||'',
      mois: mois||'', categorie: categorie||'',
      transport: parseFloat(transport)||0, repas: parseFloat(repas)||0,
      commentaire: commentaire||'', vat_lines: parsedVat, images,
      total_ttc: totals.total_ttc, total_ht: totals.total_ht,
      tva_2_6: totals.tva_2_6, tva_5_5: totals.tva_5_5, tva_10: totals.tva_10, tva_20: totals.tva_20,
    }]).select().single();
    if (error) return res.status(500).json({ error: error.message });
    return res.json(mapRow(data));
  }

  const db = readDB();
  const row = {
    id: db.nextId++,
    description: description||'', client: client||'', projet: projet||'',
    mois: mois||'', categorie: categorie||'',
    transport: parseFloat(transport)||0, repas: parseFloat(repas)||0,
    commentaire: commentaire||'', vat_lines: parsedVat, images,
    total_ttc: totals.total_ttc, total_ht: totals.total_ht,
    tva_2_6: totals.tva_2_6, tva_5_5: totals.tva_5_5, tva_10: totals.tva_10, tva_20: totals.tva_20,
    created_at: new Date().toISOString(),
  };
  db.expenses.push(row);
  writeDB(db);
  res.json(mapRow(row));
});

// ── API : modification ────────────────────────────────────
app.put('/api/expenses/:id', upload.array('images', 20), async (req, res) => {
  const id = parseInt(req.params.id);
  const { description, client, projet, mois, categorie, transport, repas, commentaire, vatLines, removeImages } = req.body;
  const toRemove  = removeImages ? JSON.parse(removeImages) : [];
  const parsedVat = JSON.parse(vatLines||'[]');
  const totals    = computeTotals(parsedVat, transport, repas);

  if (USE_SUPABASE) {
    const { data: existing } = await supabase.from('expenses').select('images').eq('id', id).single();
    if (!existing) return res.status(404).json({ error: 'Non trouvé' });
    await deleteImages(toRemove);
    const images = [...(existing.images||[]).filter(u => !toRemove.includes(u)), ...await uploadImages(req.files||[])];
    const { data, error } = await supabase.from('expenses').update({
      description: description||'', client: client||'', projet: projet||'',
      mois: mois||'', categorie: categorie||'',
      transport: parseFloat(transport)||0, repas: parseFloat(repas)||0,
      commentaire: commentaire||'', vat_lines: parsedVat, images,
      total_ttc: totals.total_ttc, total_ht: totals.total_ht,
      tva_2_6: totals.tva_2_6, tva_5_5: totals.tva_5_5, tva_10: totals.tva_10, tva_20: totals.tva_20,
    }).eq('id', id).select().single();
    if (error) return res.status(500).json({ error: error.message });
    return res.json(mapRow(data));
  }

  const db = readDB();
  const idx = db.expenses.findIndex(e => e.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Non trouvé' });
  await deleteImages(toRemove);
  const images = [...(db.expenses[idx].images||[]).filter(u => !toRemove.includes(u)), ...await uploadImages(req.files||[])];
  db.expenses[idx] = { ...db.expenses[idx], description: description||'', client: client||'', projet: projet||'', mois: mois||'', categorie: categorie||'', transport: parseFloat(transport)||0, repas: parseFloat(repas)||0, commentaire: commentaire||'', vat_lines: parsedVat, images, total_ttc: totals.total_ttc, total_ht: totals.total_ht, tva_2_6: totals.tva_2_6, tva_5_5: totals.tva_5_5, tva_10: totals.tva_10, tva_20: totals.tva_20 };
  writeDB(db);
  res.json(mapRow(db.expenses[idx]));
});

// ── API : suppression ─────────────────────────────────────
app.delete('/api/expenses/:id', async (req, res) => {
  const id = parseInt(req.params.id);
  if (USE_SUPABASE) {
    const { data: existing } = await supabase.from('expenses').select('images').eq('id', id).single();
    if (existing?.images?.length) await deleteImages(existing.images);
    await supabase.from('expenses').delete().eq('id', id);
    return res.json({ ok: true });
  }
  const db = readDB();
  const row = db.expenses.find(e => e.id === id);
  if (row?.images?.length) await deleteImages(row.images);
  db.expenses = db.expenses.filter(e => e.id !== id);
  writeDB(db);
  res.json({ ok: true });
});

// ── Requête expenses filtrée ──────────────────────────────
async function getExpenses(moisFilter) {
  if (USE_SUPABASE) {
    let q = supabase.from('expenses').select('*').order('mois').order('id');
    if (moisFilter.length) q = q.or(moisFilter.map(m => `mois.like.${m}%`).join(','));
    const { data } = await q;
    return (data||[]).map(mapRow);
  }
  const db = readDB();
  let rows = db.expenses.map(mapRow);
  if (moisFilter.length) rows = rows.filter(e => moisFilter.some(m => (e.mois||'').startsWith(m)));
  return rows.sort((a,b) => (a.mois||'').localeCompare(b.mois||'')||a.id-b.id);
}

// ── Import Excel ──────────────────────────────────────────
const uploadXlsx = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20*1024*1024 }, fileFilter: xlsxFilter });

app.post('/api/import/excel', uploadXlsx.single('file'), async (req, res) => {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.worksheets[0];

    // 1. Détecter la ligne d'en-tête et mapper les colonnes par nom
    let headerRowNum = null;
    const colMap = {};

    function cellText(cell) {
      const v = cell.value;
      if (!v) return '';
      if (typeof v === 'string') return v;
      if (typeof v === 'number') return String(v);
      if (v && v.richText) return v.richText.map(r => r.text || '').join('');
      if (v && v.result !== undefined) return String(v.result);
      return String(v);
    }

    function cellNum(cell) {
      const v = cell.value;
      if (!v && v !== 0) return 0;
      if (typeof v === 'number') return v;
      // Cellule formule : {formula, result}
      if (typeof v === 'object' && v.result !== undefined) return parseFloat(v.result) || 0;
      return parseFloat(v) || 0;
    }

    ws.eachRow((row, rowNum) => {
      if (headerRowNum) return;
      let foundTotal = false;
      row.eachCell((cell, colNum) => {
        const v = cellText(cell).replace(/\s+/g,' ').trim().toLowerCase();
        if (v === '#')                                  colMap.num        = colNum;
        else if (v.includes('description') || v.includes('client')) colMap.desc = colNum;
        else if (v === 'transport')                     colMap.transport  = colNum;
        else if (v === 'repas')                         colMap.repas      = colNum;
        else if (v === 'commentaire')                   colMap.commentaire= colNum;
        else if (v === 'total ttc')                   { colMap.totalTTC   = colNum; foundTotal = true; }
        else if (v === 'total ht')                     colMap.totalHT    = colNum;
        else if (v.includes('10%'))                    colMap.tva10      = colNum;
        else if (v.includes('20%'))                    colMap.tva20      = colNum;
        else if (v.includes('2,6%') || v.includes('2.6%')) colMap.tva26 = colNum;
        else if (v.includes('5,5%') || v.includes('5.5%')) colMap.tva55 = colNum;
      });
      if (foundTotal) headerRowNum = rowNum;
    });

    if (!headerRowNum) return res.status(400).json({ error: 'En-tête non trouvée dans le fichier.' });

    // 2. Lire les lignes de données
    const toInsert = [];

    ws.eachRow((row, rowNum) => {
      if (rowNum <= headerRowNum) return;
      const numVal = cellNum(row.getCell(colMap.num || 2));
      if (!numVal || isNaN(numVal)) return;

      const ttc = cellNum(row.getCell(colMap.totalTTC));
      if (ttc === 0) return;

      // Date : colonne C (index 3) = Description/Client/Projet
      let mois = '';
      const dateCell = row.getCell(3); // toujours col C dans ce template
      const dateVal  = dateCell.value;
      if (dateVal instanceof Date) {
        const y=dateVal.getFullYear(), m=String(dateVal.getMonth()+1).padStart(2,'0'), d=String(dateVal.getDate()).padStart(2,'0');
        mois = `${y}-${m}-${d}`;
      } else if (typeof dateVal === 'string') {
        const parts = dateVal.trim().split('/');
        if (parts.length === 3) mois = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
      }

      const transport   = colMap.transport   ? cellNum(row.getCell(colMap.transport))   : 0;
      const repas       = colMap.repas       ? cellNum(row.getCell(colMap.repas))       : 0;
      const commentaire = colMap.commentaire ? cellText(row.getCell(colMap.commentaire)).trim() : '';
      const totalHT     = colMap.totalHT     ? cellNum(row.getCell(colMap.totalHT))     : 0;
      const tva10       = colMap.tva10       ? cellNum(row.getCell(colMap.tva10))       : 0;
      const tva20       = colMap.tva20       ? cellNum(row.getCell(colMap.tva20))       : 0;
      const tva26       = colMap.tva26       ? cellNum(row.getCell(colMap.tva26))       : 0;
      const tva55       = colMap.tva55       ? cellNum(row.getCell(colMap.tva55))       : 0;

      // Reconstruire les vatLines depuis la ventilation
      const vatLines = [];
      if (tva10 && ttc) vatLines.push({ rate: 10,  ttc: transport || repas || ttc });
      if (tva20 && ttc) vatLines.push({ rate: 20,  ttc: transport || repas || ttc });
      if (tva26 && ttc) vatLines.push({ rate: 2.6, ttc: transport || repas || ttc });
      if (tva55 && ttc) vatLines.push({ rate: 5.5, ttc: transport || repas || ttc });

      const cat = transport > 0 ? 'Transport' : repas > 0 ? 'Repas' : '';

      toInsert.push({
        description: '', client: '', projet: '', commentaire,
        mois, categorie: cat,
        transport, repas,
        vat_lines: vatLines,
        total_ttc: ttc,
        total_ht:  totalHT || ttc - tva10 - tva20 - tva26 - tva55,
        tva_10: tva10, tva_20: tva20, tva_2_6: tva26, tva_5_5: tva55,
        images: [],
      });
    });

    if (!toInsert.length) return res.json({ imported: 0 });

    if (USE_SUPABASE) {
      const { error } = await supabase.from('expenses').insert(toInsert);
      if (error) return res.status(500).json({ error: error.message });
    } else {
      const db = readDB();
      toInsert.forEach(row => { db.expenses.push({ ...row, id: db.nextId++, created_at: new Date().toISOString() }); });
      writeDB(db);
    }

    res.json({ imported: toInsert.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Export Excel ──────────────────────────────────────────
app.get('/api/export/excel', async (req, res) => {
  const moisFilter = [].concat(req.query['mois[]']||req.query.mois||[]).filter(Boolean);
  const expenses   = await getExpenses(moisFilter);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'NDF App';
  const ws = wb.addWorksheet('Notes de Frais', { pageSetup: { orientation: 'landscape', fitToPage: true } });

  const BLUE = 'FF1a1a18', BL2 = 'FFf4f4f4', WHITE = 'FFffffff', DARK = 'FF0a0a0a';
  const border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

  ws.columns = [
    { header: 'ID',                 key: 'id',          width: 6  },
    { header: 'Description',        key: 'description', width: 22 },
    { header: 'Client',             key: 'client',      width: 18 },
    { header: 'Projet',             key: 'projet',      width: 18 },
    { header: 'Catégorie',          key: 'categorie',   width: 14 },
    { header: 'Date',               key: 'mois',        width: 12 },
    { header: 'Transport (TTC)',     key: 'transport',   width: 16 },
    { header: 'Repas (TTC)',         key: 'repas',       width: 14 },
    { header: 'Commentaire',        key: 'commentaire', width: 24 },
    { header: 'TOTAL TTC',          key: 'totalTTC',    width: 14 },
    { header: 'TOTAL HT',           key: 'totalHT',     width: 14 },
    { header: 'TVA 10% (Montant)',   key: 'tva_10',      width: 16 },
    { header: 'TVA 20% (Montant)',   key: 'tva_20',      width: 16 },
    { header: 'TVA 2,6% (Montant)',  key: 'tva_2_6',     width: 16 },
    { header: 'TVA 5,5% (Montant)',  key: 'tva_5_5',     width: 16 },
    { header: 'Justificatifs',      key: 'imageCount',  width: 14 },
  ];

  ws.getRow(1).eachCell(cell => {
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:BLUE} };
    cell.font = { bold:true, color:{argb:WHITE}, size:10 };
    cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
    cell.border = border;
  });
  ws.getRow(1).height = 28;
  ws.views = [{ state:'frozen', ySplit:1 }];

  const money = '#,##0.00 €';
  let sumT=0,sumR=0,sumTTC=0,sumHT=0,s10=0,s20=0,s26=0,s55=0;

  expenses.forEach((e,i) => {
    sumT+=e.transport; sumR+=e.repas; sumTTC+=e.totalTTC; sumHT+=e.totalHT;
    s10+=e.tva_10; s20+=e.tva_20; s26+=e.tva_2_6; s55+=e.tva_5_5;
    const row = ws.addRow({ id:e.id, description:e.description, client:e.client, projet:e.projet, categorie:e.categorie, mois:e.mois, transport:e.transport, repas:e.repas, commentaire:e.commentaire, totalTTC:e.totalTTC, totalHT:e.totalHT, tva_10:e.tva_10||'', tva_20:e.tva_20||'', tva_2_6:e.tva_2_6||'', tva_5_5:e.tva_5_5||'', imageCount:e.images.length });
    if (i%2===1) row.eachCell(cell => { cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:BL2}}; });
    ['transport','repas','totalTTC','totalHT','tva_10','tva_20','tva_2_6','tva_5_5'].forEach(k => { const c=row.getCell(k); if(typeof e[k]==='number'&&e[k]!==0) c.numFmt=money; });
    row.getCell('imageCount').alignment = { horizontal:'center' };
    row.eachCell(cell => { cell.border = border; });
  });

  const tr = ws.addRow({ id:'TOTAL', transport:sumT, repas:sumR, totalTTC:sumTTC, totalHT:sumHT, tva_10:s10, tva_20:s20, tva_2_6:s26, tva_5_5:s55 });
  tr.eachCell(cell => { cell.font={bold:true,color:{argb:DARK}}; cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFe8e8e4'}}; cell.border=border; });
  ['transport','repas','totalTTC','totalHT','tva_10','tva_20','tva_2_6','tva_5_5'].forEach(k => tr.getCell(k).numFmt=money);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="NDF_${new Date().toISOString().slice(0,10)}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
});

// ── Export PDF ────────────────────────────────────────────
function fetchImageBuffer(url, redirects=5) {
  return new Promise((resolve, reject) => {
    if (redirects===0) return reject(new Error('Too many redirects'));
    https.get(url, r => {
      if (r.statusCode>=300&&r.statusCode<400&&r.headers.location) return resolve(fetchImageBuffer(r.headers.location, redirects-1));
      const chunks=[];
      r.on('data',c=>chunks.push(c));
      r.on('end',()=>resolve(Buffer.concat(chunks)));
      r.on('error',reject);
    }).on('error',reject);
  });
}

async function readImageBuffer(urlOrPath) {
  if (urlOrPath.startsWith('http')) return fetchImageBuffer(urlOrPath);
  return fs.readFileSync(path.join(__dirname, urlOrPath));
}

app.get('/api/export/pdf', async (req, res) => {
  const moisFilter = [].concat(req.query['mois[]']||req.query.mois||[]).filter(Boolean);
  const expenses   = await getExpenses(moisFilter);

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="NDF_Factures_${new Date().toISOString().slice(0,10)}.pdf"`);

  const doc = new PDFDocument({ size:'A4', margin:40, info:{ Title:'Notes de Frais — Justificatifs' } });
  doc.pipe(res);
  const W = doc.page.width;

  doc.rect(0,0,W,110).fill('#1a1a18');
  doc.fillColor('white').fontSize(24).font('Helvetica-Bold').text('Notes de Frais', 40, 30, {align:'center'});
  doc.fontSize(12).font('Helvetica').text(`Justificatifs — ${new Date().toLocaleDateString('fr-FR')}`, 40, 62, {align:'center'});
  doc.fontSize(10).text(`${expenses.length} dépense(s)`, 40, 84, {align:'center'});

  let y = 140;
  doc.fillColor('#0a0a0a').fontSize(13).font('Helvetica-Bold').text('Récapitulatif', 40, y);
  y += 22;

  const C = { id:{x:35,w:25}, desc:{x:63,w:140}, cat:{x:206,w:60}, mois:{x:269,w:45}, montant:{x:317,w:55}, ttc:{x:375,w:50}, ht:{x:428,w:50}, tva:{x:481,w:50} };
  const ROW_H = 16;
  doc.rect(35,y-2,W-70,ROW_H).fill('#1a1a18');
  doc.fillColor('white').fontSize(7.5).font('Helvetica-Bold');
  doc.text('ID',        C.id.x,      y, {width:C.id.w});
  doc.text('Description',C.desc.x,  y, {width:C.desc.w});
  doc.text('Catégorie', C.cat.x,    y, {width:C.cat.w});
  doc.text('Date',      C.mois.x,   y, {width:C.mois.w});
  doc.text('Montant',   C.montant.x,y, {width:C.montant.w, align:'right'});
  doc.text('TTC',       C.ttc.x,    y, {width:C.ttc.w,     align:'right'});
  doc.text('HT',        C.ht.x,     y, {width:C.ht.w,      align:'right'});
  doc.text('TVA',       C.tva.x,    y, {width:C.tva.w,     align:'right'});

  let sumTTC=0, sumHT=0;
  expenses.forEach((e,i) => {
    y += ROW_H;
    if (y>doc.page.height-60) { doc.addPage(); y=40; }
    if (i%2===0) doc.rect(35,y-2,W-70,ROW_H).fill('#f4f4f4');
    doc.fillColor('#0a0a0a').fontSize(7).font('Helvetica');
    const label   = [e.description,e.client].filter(Boolean).join(' / ');
    const montant = (e.transport||e.repas||0).toFixed(2);
    const tva     = (e.tva_10+e.tva_20+e.tva_2_6+e.tva_5_5).toFixed(2);
    doc.text(String(e.id),               C.id.x,      y, {width:C.id.w,      lineBreak:false});
    doc.text(label||'—',                 C.desc.x,    y, {width:C.desc.w,    lineBreak:false});
    doc.text(e.categorie||'—',           C.cat.x,     y, {width:C.cat.w,     lineBreak:false});
    doc.text(e.mois||'—',                C.mois.x,    y, {width:C.mois.w,    lineBreak:false});
    doc.text(montant+' €',               C.montant.x, y, {width:C.montant.w, align:'right', lineBreak:false});
    doc.text(e.totalTTC.toFixed(2)+' €', C.ttc.x,     y, {width:C.ttc.w,    align:'right', lineBreak:false});
    doc.text(e.totalHT.toFixed(2)+' €',  C.ht.x,      y, {width:C.ht.w,     align:'right', lineBreak:false});
    doc.text(tva+' €',                   C.tva.x,     y, {width:C.tva.w,     align:'right', lineBreak:false});
    sumTTC+=e.totalTTC; sumHT+=e.totalHT;
  });

  y += ROW_H+4;
  doc.rect(35,y-2,W-70,ROW_H).fill('#e8e8e4');
  doc.fillColor('#0a0a0a').fontSize(8).font('Helvetica-Bold');
  doc.text('TOTAL',                  C.id.x,  y, {width:200,        lineBreak:false});
  doc.text(sumTTC.toFixed(2)+' €',   C.ttc.x, y, {width:C.ttc.w,   align:'right', lineBreak:false});
  doc.text(sumHT.toFixed(2)+' €',    C.ht.x,  y, {width:C.ht.w,    align:'right', lineBreak:false});

  for (const e of expenses) {
    if (!e.images.length) continue;
    doc.addPage();
    doc.rect(0,0,W,80).fill('#1a1a18');
    doc.fillColor('#aaa').fontSize(9).font('Helvetica').text(`NDF #${e.id}`, 40, 12, {width:W-80});
    doc.fillColor('white').fontSize(13).font('Helvetica-Bold');
    const label = [e.description,e.client,e.projet].filter(Boolean).join(' — ')||`NDF #${e.id}`;
    doc.text(label, 40, 26, {width:W-80});
    doc.fontSize(9).font('Helvetica');
    const catLabel = e.categorie ? `${e.categorie}: ${(e.transport||e.repas||0).toFixed(2)} €` : `Transport: ${e.transport.toFixed(2)} €  |  Repas: ${e.repas.toFixed(2)} €`;
    doc.text(`${catLabel}  |  TTC: ${e.totalTTC.toFixed(2)} €  |  HT: ${e.totalHT.toFixed(2)} €`, 40, 48, {width:W-80});
    const parts=[];
    if(e.tva_2_6) parts.push(`TVA 2,6%: ${e.tva_2_6.toFixed(2)} €`);
    if(e.tva_5_5) parts.push(`TVA 5,5%: ${e.tva_5_5.toFixed(2)} €`);
    if(e.tva_10)  parts.push(`TVA 10%: ${e.tva_10.toFixed(2)} €`);
    if(e.tva_20)  parts.push(`TVA 20%: ${e.tva_20.toFixed(2)} €`);
    if(parts.length) doc.text(parts.join('  |  '), 40, 62, {width:W-80});

    let imgY=95;
    const maxH = (doc.page.height-imgY-40)/Math.min(e.images.length,2);
    for (let i=0; i<e.images.length; i++) {
      if (i>0&&i%2===0) { doc.addPage(); imgY=40; }
      try {
        const buf = await readImageBuffer(e.images[i]);
        doc.image(buf, 40, imgY, { fit:[W-80, maxH-10], align:'center' });
        imgY+=maxH;
      } catch(_) {
        doc.fillColor('#999').fontSize(9).text('[Image non disponible]', 40, imgY);
        imgY+=30;
      }
    }
  }
  doc.end();
});

app.listen(PORT, () => console.log(`NDF App : http://localhost:${PORT}`));
module.exports = app;
