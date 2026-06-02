import XLSX from 'xlsx';
import { TEMPLATE_B64 } from '../../lib/orderTemplate';

// 사이즈 → 규격 표기
const SIZE_FORMAT = {
  M1: 'M1(빨강)',
  M2: 'M2(노랑)',
  L:  'L(주황)',
  XL: 'XL(검정)',
};

// 제품 → 품명
const PRODUCT_NAME = {
  mediquet: 'Mediquet',
  rapband:  'Rap Band',
};

// Excel 날짜 시리얼 번호 계산
function excelDateNum(date) {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return (date.getTime() - epoch.getTime()) / 86400000;
}

export default function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { orders } = req.body;
  // orders: [{ product: 'mediquet'|'rapband', size: 'M1'|..., qty: 60 }]

  const wb = XLSX.read(TEMPLATE_B64, { type: 'base64' });
  const ws = wb.Sheets['렙메디케어'];

  // 오늘 날짜 설정 (D4)
  const today = new Date();
  ws['D4'] = {
    t: 'n',
    v: excelDateNum(today),
    z: 'yyyy"년" m"월" d"일"',
    w: `${today.getFullYear()}년 ${today.getMonth() + 1}월 ${today.getDate()}일`,
  };

  // 기존 주문 행 초기화 (10~22)
  for (let row = 10; row <= 22; row++) {
    ['A', 'B', 'D', 'F', 'H', 'K', 'M', 'P', 'S'].forEach((col) => {
      if (ws[`${col}${row}`]) delete ws[`${col}${row}`];
    });
  }

  // 주문 데이터 채우기
  orders.forEach((item, idx) => {
    const row = 10 + idx;
    if (row > 22) return; // 템플릿 최대 행 초과 방지

    ws[`A${row}`] = { t: 'n', v: idx + 1 };
    ws[`D${row}`] = { t: 's', v: PRODUCT_NAME[item.product] ?? item.product };
    ws[`F${row}`] = { t: 's', v: SIZE_FORMAT[item.size] ?? item.size };
    ws[`H${row}`] = { t: 'n', v: item.qty }; // 사용분(발주수량)
  });

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const y = today.getFullYear();
  const m = String(today.getMonth() + 1).padStart(2, '0');
  const d = String(today.getDate()).padStart(2, '0');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%EB%B0%9C%EC%A3%BC%EC%84%9C_${y}${m}${d}.xlsx`);
  res.send(buf);
}
