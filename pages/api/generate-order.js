import ExcelJS from 'exceljs';
import { TEMPLATE_B64 } from '../../lib/orderTemplate';
import { MEDYSSEY_B64 } from '../../lib/medysseyTemplate';
import { VENDOR_MAP, TEMPLATES } from '../../lib/vendors';

const B64_MAP = {
  repmedicare: TEMPLATE_B64,
  medyssey: MEDYSSEY_B64,
};

function styleDataCell(cell) {
  cell.font = { ...(cell.font ?? {}), name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
}

function buildRows(items, groupGap) {
  const rows = [];
  let num = 1;
  let prevName = null;
  for (const item of items) {
    if (groupGap && prevName !== null && item.name !== prevName) {
      rows.push(null); // blank gap row
    }
    rows.push({ ...item, num: num++ });
    prevName = item.name;
  }
  return rows;
}

function fillSheet(ws, rows, template, vendor, requester, cols, note) {
  const { startRow, endRow } = template.itemTable;
  const { date: dateCell, requester: reqCell, recipient: recCell, note: noteCell } = template.cells;
  const maxRows = endRow - startRow + 1;

  if (dateCell) ws.getCell(dateCell).value = new Date();
  if (reqCell && requester) ws.getCell(reqCell).value = requester;
  if (recCell) ws.getCell(recCell).value = vendor.name;
  if (noteCell && note) ws.getCell(noteCell).value = note;

  // Clear item rows
  for (let r = startRow; r <= endRow; r++) {
    ws.getCell(`${cols.no}${r}`).value = null;
    ws.getCell(`${cols.name}${r}`).value = null;
    ws.getCell(`${cols.spec}${r}`).value = null;
    ws.getCell(`${cols.qty}${r}`).value = null;
    if (cols.note) ws.getCell(`${cols.note}${r}`).value = null;
  }

  // Fill rows
  rows.slice(0, maxRows).forEach((row, i) => {
    const r = startRow + i;
    if (!row) return; // blank gap
    ws.getCell(`${cols.no}${r}`).value = row.num;
    const nameCell = ws.getCell(`${cols.name}${r}`);
    nameCell.value = row.name;
    styleDataCell(nameCell);
    const specCell = ws.getCell(`${cols.spec}${r}`);
    specCell.value = row.spec || null;
    styleDataCell(specCell);
    const qtyCell = ws.getCell(`${cols.qty}${r}`);
    qtyCell.value = row.qty;
    styleDataCell(qtyCell);
    if (cols.note && row.note) {
      const rowNoteCell = ws.getCell(`${cols.note}${r}`);
      rowNoteCell.value = row.note;
      styleDataCell(rowNoteCell);
    }
  });
}

async function cloneSheet(templateBuffer, wb, srcName, newName) {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(templateBuffer);
  const src = wb2.getWorksheet(srcName) || wb2.worksheets[0];

  const dst = wb.addWorksheet(newName, { pageSetup: { ...src.pageSetup } });

  // Copy column widths
  src.columns.forEach((col) => {
    if (col.number && col.width) dst.getColumn(col.number).width = col.width;
  });

  // Copy merges
  (src.model.merges || []).forEach((m) => {
    try { dst.mergeCells(m); } catch {}
  });

  // Copy cells (style + value)
  src.eachRow({ includeEmpty: true }, (row, rn) => {
    const dstRow = dst.getRow(rn);
    dstRow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      const dc = dstRow.getCell(cn);
      if (cell.style) dc.style = JSON.parse(JSON.stringify(cell.style));
      dc.value = cell.value;
    });
    dstRow.commit();
  });

  return dst;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { vendorId, requester, items, hospital, note } = req.body || {};

  // Support legacy format: { orders: [{product, size, qty}] }
  if (!vendorId && req.body?.orders) {
    return legacyHandler(req, res);
  }

  const vendor = VENDOR_MAP[vendorId];
  if (!vendor) return res.status(400).json({ error: 'Unknown vendorId' });

  const template = TEMPLATES[vendor.templateId];
  const b64 = B64_MAP[vendor.templateId];
  if (!template || !b64) return res.status(400).json({ error: 'Unknown template' });

  const validItems = (items || []).filter((it) => it.qty > 0);
  if (!validItems.length) return res.status(400).json({ error: 'No items' });

  const { cols, startRow, endRow, groupGap } = template.itemTable;
  const sheetName = template.sheetName;
  const maxPerPage = endRow - startRow + 1;

  const allRows = buildRows(validItems, groupGap);

  // Chunk into pages
  const pages = [];
  let page = [];
  for (const row of allRows) {
    if (page.length >= maxPerPage) {
      pages.push(page);
      page = [];
    }
    page.push(row);
  }
  if (page.length) pages.push(page);

  const templateBuffer = Buffer.from(b64, 'base64');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuffer);

  const ws1 = wb.getWorksheet(sheetName) || wb.worksheets[0];
  fillSheet(ws1, pages[0] || [], template, vendor, requester, cols, note);

  for (let p = 1; p < pages.length; p++) {
    const wsN = await cloneSheet(templateBuffer, wb, sheetName, `${sheetName} (${p + 1})`);
    fillSheet(wsN, pages[p], template, vendor, requester, cols, note);
  }

  const buffer = await wb.xlsx.writeBuffer();
  const today = new Date();
  const ds = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
  const hospPart = hospital ? encodeURIComponent(hospital) + '_' : '';
  const encVendor = encodeURIComponent(vendor.name);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${hospPart}${encVendor}_%EB%B0%9C%EC%A3%BC%EC%84%9C_${ds}.xlsx`);
  res.send(buffer);
}

// Legacy handler for old format: { orders: [{product, size, qty}] }
async function legacyHandler(req, res) {
  const { TEMPLATE_B64: tmplB64 } = await import('../../lib/orderTemplate');
  const { PRODUCT_MAP } = await import('../../lib/products');
  const { orders } = req.body;

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(tmplB64, 'base64'));
  const ws = wb.getWorksheet('렙메디케어');

  ws.getCell('D4').value = new Date();
  for (let r = 10; r <= 22; r++) {
    ['A', 'D', 'F', 'H', 'P'].forEach((c) => { ws.getCell(`${c}${r}`).value = null; });
  }

  const groups = [];
  let cur = null;
  for (const item of orders) {
    if (item.product !== cur) { groups.push([]); cur = item.product; }
    groups[groups.length - 1].push(item);
  }

  let offset = 0;
  let num = 1;
  groups.forEach((group, gi) => {
    if (gi > 0) offset++;
    group.forEach((item) => {
      const r = 10 + offset;
      if (r > 22) return;
      ws.getCell(`A${r}`).value = num++;
      const dc = ws.getCell(`D${r}`);
      dc.value = PRODUCT_MAP[item.product]?.excelName ?? item.product;
      dc.font = { name: '맑은 고딕' };
      dc.alignment = { horizontal: 'center', vertical: 'middle' };
      const fc = ws.getCell(`F${r}`);
      const sizeObj = PRODUCT_MAP[item.product]?.sizes.find((s) => s.id === item.size);
      fc.value = sizeObj?.excelSize ?? item.size;
      fc.font = { name: '맑은 고딕' };
      fc.alignment = { horizontal: 'center', vertical: 'middle' };
      const hc = ws.getCell(`H${r}`);
      hc.value = item.qty;
      hc.font = { name: '맑은 고딕' };
      hc.alignment = { horizontal: 'center', vertical: 'middle' };
      offset++;
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  const today = new Date();
  const ds = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%EB%B0%9C%EC%A3%BC%EC%84%9C_${ds}.xlsx`);
  res.send(buffer);
}
