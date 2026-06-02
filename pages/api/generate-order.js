import ExcelJS from 'exceljs';
import { TEMPLATE_B64 } from '../../lib/orderTemplate';
import { PRODUCT_MAP } from '../../lib/products';

function getSizeExcel(productId, sizeId) {
  const sizeObj = PRODUCT_MAP[productId]?.sizes.find(s => s.id === sizeId);
  return sizeObj?.excelSize ?? sizeId;
}

function styleDataCell(cell) {
  cell.font = { ...(cell.font ?? {}), name: '맑은 고딕' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { orders } = req.body;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(TEMPLATE_B64, 'base64'));
  const ws = workbook.getWorksheet('렙메디케어');

  ws.getCell('D4').value = new Date();

  for (let row = 10; row <= 22; row++) {
    ws.getCell(`A${row}`).value = null;
    ws.getCell(`D${row}`).value = null;
    ws.getCell(`F${row}`).value = null;
    ws.getCell(`H${row}`).value = null;
    ws.getCell(`P${row}`).value = null;
  }

  // Group by product, insert a blank row between groups
  const groups = [];
  let curProd = null;
  for (const item of orders) {
    if (item.product !== curProd) { groups.push([]); curProd = item.product; }
    groups[groups.length - 1].push(item);
  }

  let rowOffset = 0;
  let orderNum  = 1;

  groups.forEach((group, gi) => {
    if (gi > 0) rowOffset++;
    group.forEach((item) => {
      const row = 10 + rowOffset;
      if (row > 22) return;

      ws.getCell(`A${row}`).value = orderNum++;

      const dCell = ws.getCell(`D${row}`);
      dCell.value = PRODUCT_MAP[item.product]?.excelName ?? item.product;
      styleDataCell(dCell);

      const fCell = ws.getCell(`F${row}`);
      fCell.value = getSizeExcel(item.product, item.size);
      styleDataCell(fCell);

      const hCell = ws.getCell(`H${row}`);
      hCell.value = item.qty;
      styleDataCell(hCell);

      rowOffset++;
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const today  = new Date();
  const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%EB%B0%9C%EC%A3%BC%EC%84%9C_${ds}.xlsx`);
  res.send(buffer);
}
