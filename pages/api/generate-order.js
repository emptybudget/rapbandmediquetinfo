import ExcelJS from 'exceljs';
import { TEMPLATE_B64 } from '../../lib/orderTemplate';

const SIZE_FORMAT = {
  M1: 'M1(빨강)',
  M2: 'M2(노랑)',
  L:  'L(주황)',
  XL: 'XL(검정)',
};

const PRODUCT_NAME = {
  mediquet: 'Mediquet',
  rapband:  'Rap Band',
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { orders } = req.body;

  // 원본 템플릿 로드 (ExcelJS는 서식·병합 셀 완전 보존)
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(TEMPLATE_B64, 'base64'));

  const ws = workbook.getWorksheet('렙메디케어');

  // 날짜만 오늘로 변경 (D4) — 셀 서식 유지
  ws.getCell('D4').value = new Date();

  // 기존 주문행 값 초기화 (셀 서식·병합은 그대로)
  for (let row = 10; row <= 22; row++) {
    ws.getCell(`A${row}`).value = null;
    ws.getCell(`D${row}`).value = null;
    ws.getCell(`F${row}`).value = null;
    ws.getCell(`H${row}`).value = null;
    ws.getCell(`P${row}`).value = null;
  }

  // 발주 데이터 기입
  orders.forEach((item, idx) => {
    const row = 10 + idx;
    if (row > 22) return;
    ws.getCell(`A${row}`).value = idx + 1;
    ws.getCell(`D${row}`).value = PRODUCT_NAME[item.product] ?? item.product;
    ws.getCell(`F${row}`).value = SIZE_FORMAT[item.size] ?? item.size;
    ws.getCell(`H${row}`).value = item.qty; // 사용분(발주수량)
  });

  const buffer = await workbook.xlsx.writeBuffer();

  const today = new Date();
  const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%EB%B0%9C%EC%A3%BC%EC%84%9C_${ds}.xlsx`);
  res.send(buffer);
}
