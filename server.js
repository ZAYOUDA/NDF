require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { createClient } = require('@supabase/supabase-js');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const https = require('https');

const app  = express();
const PORT = process.env.PORT || 3000;

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

const BUCKET = 'receipts';

// multer — mémoire uniquement (pas de disque)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Seules les images sont acceptées'));
  },
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Helpers ──────────────────────────────────────────────
function computeTotals(vatLines) {
  const r = { tva_2_6: 0, tva_5_5: 0, tva_10: 0, tva_20: 0, total_ht: 0, total_ttc: 0 };
  (vatLines || []).forEach(line => {
    const ttc  = parseFloat(line.ttc)  || 0;
    const rate = parseFloat(line.rate) || 0;
    const ht   = ttc / (1 + rate / 100);
    r.total_ttc += ttc;
    r.total_ht  += ht;
    if (rate === 2.6)  r.tva_2_6 += ttc - ht;
    if (rate === 5.5)  r.tva_5_5 += ttc - ht;
    if (rate === 10)   r.tva_10  += ttc - ht;
    if (rate === 20)   r.tva_20  += ttc - ht;
  });
  return r;
}

async function uploadImages(files) {
  const urls = [];
  for (const file of files) {
    const ext      = path.extname(file.originalname) || '.jpg';
    const filename = `${uuidv4()}${ext}`;
    const { error } = await supabase.storage
      .from(BUCKET)
      .upload(filename, file.buffer, { contentType: file.mimetype, upsert: false });
    if (error) throw error;
    const { data } = supabase.storage.from(BUCKET).getPublicUrl(filename);
    urls.push(data.publicUrl);
  }
  return urls;
}

async function deleteImages(urls) {
  const paths = urls.map(u => u.split(`${BUCKET}/`)[1]).filter(Boolean);
  if (paths.length) await supabase.storage.from(BUCKET).remove(paths);
}

function mapRow(row) {
  return {
    id:          row.id,
    description: row.description || '',
    client:      row.client      || '',
    projet:      row.projet      || '',
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
    createdAt:   row.created_at,
  };
}

// ── API : liste ───────────────────────────────────────────
app.get('/api/expenses', async (req, res) => {
  const { data, error } = await supabase
    .from('expenses')
    .select('*')
    .order('id', { ascending: true });
  if (error) return res.status(500).json({ error: error.message });
  res.json(data.map(mapRow));
});

// ── API : ajout ───────────────────────────────────────────
app.post('/api/expenses', upload.array('images', 20), async (req, res) => {
  const { description, client, projet, transport, repas, commentaire, vatLines } = req.body;
  const parsedVat = JSON.parse(vatLines || '[]');
  const totals    = computeTotals(parsedVat);
  const images    = await uploadImages(req.files || []);

  const { data, error } = await supabase.from('expenses').insert([{
    description:  description  || '',
    client:       client       || '',
    projet:       projet       || '',
    transport:    parseFloat(transport) || 0,
    repas:        parseFloat(repas)     || 0,
    commentaire:  commentaire  || '',
    vat_lines:    parsedVat,
    total_ttc:    totals.total_ttc,
    total_ht:     totals.total_ht,
    tva_2_6:      totals.tva_2_6,
    tva_5_5:      totals.tva_5_5,
    tva_10:       totals.tva_10,
    tva_20:       totals.tva_20,
    images,
  }]).select().single();

  if (error) return res.status(500).json({ error: error.message });
  res.json(mapRow(data));
});

// ── API : modification ────────────────────────────────────
app.put('/api/expenses/:id', upload.array('images', 20), async (req, res) => {
  const id = parseInt(req.params.id);
  const { description, client, projet, transport, repas, commentaire, vatLines, removeImages } = req.body;

  const { data: existing } = await supabase.from('expenses').select('images').eq('id', id).single();
  if (!existing) return res.status(404).json({ error: 'Non trouvé' });

  const toRemove = removeImages ? JSON.parse(removeImages) : [];
  await deleteImages(toRemove);

  const keptImages = (existing.images || []).filter(u => !toRemove.includes(u));
  const newImages  = await uploadImages(req.files || []);
  const images     = [...keptImages, ...newImages];

  const parsedVat = JSON.parse(vatLines || '[]');
  const totals    = computeTotals(parsedVat);

  const { data, error } = await supabase.from('expenses').update({
    description:  description  || '',
    client:       client       || '',
    projet:       projet       || '',
    transport:    parseFloat(transport) || 0,
    repas:        parseFloat(repas)     || 0,
    commentaire:  commentaire  || '',
    vat_lines:    parsedVat,
    total_ttc:    totals.total_ttc,
    total_ht:     totals.total_ht,
    tva_2_6:      totals.tva_2_6,
    tva_5_5:      totals.tva_5_5,
    tva_10:       totals.tva_10,
    tva_20:       totals.tva_20,
    images,
  }).eq('id', id).select().single();

  if (error) return res.status(500).json({ error: error.message });
  res.json(mapRow(data));
});

// ── API : suppression ─────────────────────────────────────
app.delete('/api/expenses/:id', async (req, res) => {
  const id = parseInt(req.params.id);
  const { data: existing } = await supabase.from('expenses').select('images').eq('id', id).single();
  if (existing?.images?.length) await deleteImages(existing.images);
  await supabase.from('expenses').delete().eq('id', id);
  res.json({ ok: true });
});

// ── Export Excel ──────────────────────────────────────────
app.get('/api/export/excel', async (req, res) => {
  const { data: rows } = await supabase.from('expenses').select('*').order('id');
  const expenses = (rows || []).map(mapRow);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'NDF App';
  const ws = wb.addWorksheet('Notes de Frais', { pageSetup: { orientation: 'landscape', fitToPage: true } });

  const BLUE = 'FF1a1a18', BL2 = 'FFf4f4f4', WHITE = 'FFffffff', DARK = 'FF0a0a0a';
  const border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

  ws.columns = [
    { header: 'ID',                  key: 'id',          width: 6  },
    { header: 'Description',         key: 'description', width: 22 },
    { header: 'Client',              key: 'client',      width: 18 },
    { header: 'Projet',              key: 'projet',      width: 18 },
    { header: 'Transport (TTC)',      key: 'transport',   width: 16 },
    { header: 'Repas (TTC)',          key: 'repas',       width: 14 },
    { header: 'Commentaire',         key: 'commentaire', width: 24 },
    { header: 'TOTAL TTC',           key: 'totalTTC',    width: 14 },
    { header: 'TOTAL HT',            key: 'totalHT',     width: 14 },
    { header: 'TVA 10% (Montant)',    key: 'tva_10',      width: 16 },
    { header: 'TVA 20% (Montant)',    key: 'tva_20',      width: 16 },
    { header: 'TVA 2,6% (Montant)',   key: 'tva_2_6',     width: 16 },
    { header: 'TVA 5,5% (Montant)',   key: 'tva_5_5',     width: 16 },
    { header: 'Justificatifs',       key: 'imageCount',  width: 14 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
    cell.font   = { bold: true, color: { argb: WHITE }, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = border;
  });
  headerRow.height = 28;
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const money = '#,##0.00 €';
  let sumT=0,sumR=0,sumTTC=0,sumHT=0,s10=0,s20=0,s26=0,s55=0;

  expenses.forEach((e, i) => {
    sumT+=e.transport; sumR+=e.repas; sumTTC+=e.totalTTC; sumHT+=e.totalHT;
    s10+=e.tva_10; s20+=e.tva_20; s26+=e.tva_2_6; s55+=e.tva_5_5;

    const row = ws.addRow({
      id: e.id, description: e.description, client: e.client, projet: e.projet,
      transport: e.transport, repas: e.repas, commentaire: e.commentaire,
      totalTTC: e.totalTTC, totalHT: e.totalHT,
      tva_10: e.tva_10||'', tva_20: e.tva_20||'', tva_2_6: e.tva_2_6||'', tva_5_5: e.tva_5_5||'',
      imageCount: e.images.length,
    });

    if (i % 2 === 1) row.eachCell(cell => { cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:BL2} }; });
    ['transport','repas','totalTTC','totalHT','tva_10','tva_20','tva_2_6','tva_5_5'].forEach(k => {
      const c = row.getCell(k);
      if (typeof e[k] === 'number' && e[k] !== 0) c.numFmt = money;
    });
    row.getCell('imageCount').alignment = { horizontal: 'center' };
    row.eachCell(cell => { cell.border = border; });
  });

  const tr = ws.addRow({ id:'TOTAL', transport:sumT, repas:sumR, totalTTC:sumTTC, totalHT:sumHT, tva_10:s10, tva_20:s20, tva_2_6:s26, tva_5_5:s55 });
  tr.eachCell(cell => {
    cell.font = { bold:true, color:{argb:DARK} };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFe8e8e4'} };
    cell.border = border;
  });
  ['transport','repas','totalTTC','totalHT','tva_10','tva_20','tva_2_6','tva_5_5'].forEach(k => tr.getCell(k).numFmt = money);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="NDF_${new Date().toISOString().slice(0,10)}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
});

// ── Export PDF ────────────────────────────────────────────
function fetchImageBuffer(url) {
  return new Promise((resolve, reject) => {
    https.get(url, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

app.get('/api/export/pdf', async (req, res) => {
  const { data: rows } = await supabase.from('expenses').select('*').order('id');
  const expenses = (rows || []).map(mapRow);

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

  const cols = { id:40, desc:65, transport:220, repas:285, ttc:345, ht:400, tva:455 };
  doc.rect(35, y-3, W-70, 18).fill('#1a1a18');
  doc.fillColor('white').fontSize(8).font('Helvetica-Bold');
  doc.text('ID', cols.id, y, {width:22});
  doc.text('Desc / Client / Projet', cols.desc, y, {width:152});
  doc.text('Transport', cols.transport, y, {width:60, align:'right'});
  doc.text('Repas', cols.repas, y, {width:55, align:'right'});
  doc.text('TTC', cols.ttc, y, {width:50, align:'right'});
  doc.text('HT', cols.ht, y, {width:50, align:'right'});
  doc.text('TVA', cols.tva, y, {width:50, align:'right'});

  let sumTTC=0, sumHT=0;
  expenses.forEach((e,i) => {
    y += 18;
    if (y > doc.page.height-60) { doc.addPage(); y=40; }
    if (i%2===0) doc.rect(35, y-3, W-70, 16).fill('#f4f4f4');
    doc.fillColor('#0a0a0a').fontSize(7.5).font('Helvetica');
    const label = [e.description, e.client, e.projet].filter(Boolean).join(' / ');
    doc.text(String(e.id), cols.id, y, {width:22});
    doc.text(label||'—', cols.desc, y, {width:152});
    doc.text(e.transport?e.transport.toFixed(2):'—', cols.transport, y, {width:60, align:'right'});
    doc.text(e.repas?e.repas.toFixed(2):'—', cols.repas, y, {width:55, align:'right'});
    doc.text(e.totalTTC.toFixed(2), cols.ttc, y, {width:50, align:'right'});
    doc.text(e.totalHT.toFixed(2), cols.ht, y, {width:50, align:'right'});
    const tva = e.tva_10+e.tva_20+e.tva_2_6+e.tva_5_5;
    doc.text(tva.toFixed(2), cols.tva, y, {width:50, align:'right'});
    sumTTC+=e.totalTTC; sumHT+=e.totalHT;
  });

  y+=20;
  doc.rect(35, y-3, W-70, 18).fill('#e8e8e4');
  doc.fillColor('#0a0a0a').fontSize(8.5).font('Helvetica-Bold');
  doc.text('TOTAL', cols.id, y, {width:180});
  doc.text(sumTTC.toFixed(2)+' €', cols.ttc, y, {width:50, align:'right'});
  doc.text(sumHT.toFixed(2)+' €', cols.ht, y, {width:50, align:'right'});

  for (const e of expenses) {
    if (!e.images.length) continue;
    doc.addPage();
    doc.rect(0,0,W,75).fill('#1a1a18');
    doc.fillColor('white').fontSize(13).font('Helvetica-Bold');
    const label = [e.description, e.client, e.projet].filter(Boolean).join(' — ') || `NDF #${e.id}`;
    doc.text(label, 40, 12, {width:W-80});
    doc.fontSize(9).font('Helvetica');
    doc.text(`TTC: ${e.totalTTC.toFixed(2)} €  |  HT: ${e.totalHT.toFixed(2)} €  |  Transport: ${e.transport.toFixed(2)} €  |  Repas: ${e.repas.toFixed(2)} €`, 40, 36, {width:W-80});
    const parts=[];
    if(e.tva_2_6) parts.push(`TVA 2,6%: ${e.tva_2_6.toFixed(2)} €`);
    if(e.tva_5_5) parts.push(`TVA 5,5%: ${e.tva_5_5.toFixed(2)} €`);
    if(e.tva_10)  parts.push(`TVA 10%: ${e.tva_10.toFixed(2)} €`);
    if(e.tva_20)  parts.push(`TVA 20%: ${e.tva_20.toFixed(2)} €`);
    doc.text(parts.join('  |  '), 40, 53, {width:W-80});

    let imgY=90;
    const maxH = (doc.page.height-imgY-40) / Math.min(e.images.length,2);

    for (let i=0; i<e.images.length; i++) {
      if (i>0 && i%2===0) { doc.addPage(); imgY=40; }
      try {
        const buf = await fetchImageBuffer(e.images[i]);
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
