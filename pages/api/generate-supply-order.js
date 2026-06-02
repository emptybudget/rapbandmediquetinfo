import ExcelJS from 'exceljs';
import { TEMPLATE_B64 } from '../../lib/orderTemplate';

function styleDataCell(cell) {
  cell.font = { ...(cell.font ?? {}), name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { requester, recipient, items } = req.body;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(TEMPLATE_B64, 'base64'));
  const ws = workbook.getWorksheet('렙메디케어');

  ws.getCell('D4').value = new Date();
  if (requester) ws.getCell('M4').value = requester;
  if (recipient) ws.getCell('D6').value = recipient;

  // Clear rows 10–22
  for (let row = 10; row <= 22; row++) {
    ws.getCell(`A${row}`).value = null;
    ws.getCell(`D${row}`).value = null;
    ws.getCell(`F${row}`).value = null;
    ws.getCell(`H${row}`).value = null;
    ws.getCell(`P${row}`).value = null;
  }

  // Write items (max 13 rows)
  items.slice(0, 13).forEach((item, i) => {
    const row = 10 + i;
    ws.getCell(`A${row}`).value = i + 1;

    const dCell = ws.getCell(`D${row}`);
    dCell.value = item.name;
    styleDataCell(dCell);

    const fCell = ws.getCell(`F${row}`);
    fCell.value = item.spec || '';
    styleDataCell(fCell);

    const hCell = ws.getCell(`H${row}`);
    hCell.value = Number(item.qty) || 0;
    styleDataCell(hCell);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const today = new Date();
  const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  const firstName = encodeURIComponent(items[0]?.name || '소모품');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${firstName}_%EB%B0%9C%EC%A3%BC%EC%84%9C_${ds}.xlsx`);
  res.send(buffer);
}
