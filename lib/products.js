// Central product registry.
// To add a new product: append an entry to PRODUCTS below.
// type 'calculated' → stock-based order calc + optional Excel upload.
// type 'manual'     → user types the order quantity directly.

export const PACK_SIZE = 30;
export const FREE_SHIP_BOXES = 6;

export const PRODUCTS = [
  {
    id: 'mediquet',
    name: 'Mediquet',
    type: 'calculated',
    color: '#e74c3c',
    excelName: 'Mediquet',
    hasExcelUpload: true,
    defaultOcha: { M1: 22, M2: 35, L: -7, XL: -7 },
    sizes: [
      {
        id: 'M1', label: 'M1 빨강', chipColor: '#e74c3c', excelSize: 'M1(빨강)',
        avg_monthly: 36.3, monthly: { '2026-03': 17, '2026-04': 44, '2026-05': 48 }, default_stock: 52,
      },
      {
        id: 'M2', label: 'M2 노랑', chipColor: '#e6b800', excelSize: 'M2(노랑)',
        avg_monthly: 90.7, monthly: { '2026-03': 105, '2026-04': 87, '2026-05': 80 }, default_stock: 73,
      },
      {
        id: 'L', label: 'L 주황', chipColor: '#e67e22', excelSize: 'L(주황)',
        avg_monthly: 82.7, monthly: { '2026-03': 73, '2026-04': 93, '2026-05': 82 }, default_stock: 21,
      },
      {
        id: 'XL', label: 'XL 검정', chipColor: '#2c3e50', excelSize: 'XL(검정)',
        avg_monthly: 67.3, monthly: { '2026-03': 49, '2026-04': 75, '2026-05': 78 }, default_stock: 96,
      },
    ],
  },
  {
    id: 'rapband',
    name: 'Rapband',
    type: 'calculated',
    color: '#2980b9',
    excelName: 'Rapband',
    hasExcelUpload: true,
    defaultOcha: { M1: 3, M2: 1, L: -8, XL: 30 },
    sizes: [
      {
        id: 'M1', label: 'M1 빨강', chipColor: '#e74c3c', excelSize: 'M1(빨강)',
        avg_monthly: 80.7, monthly: { '2026-03': 56, '2026-04': 111, '2026-05': 75 }, default_stock: 84,
      },
      {
        id: 'M2', label: 'M2 노랑', chipColor: '#e6b800', excelSize: 'M2(노랑)',
        avg_monthly: 98.3, monthly: { '2026-03': 87, '2026-04': 123, '2026-05': 55 }, default_stock: 111,
      },
      {
        id: 'L', label: 'L 주황', chipColor: '#e67e22', excelSize: 'L(주황)',
        avg_monthly: 327.7, monthly: { '2026-03': 326, '2026-04': 318, '2026-05': 250 }, default_stock: 110,
      },
      {
        id: 'XL', label: 'XL 검정', chipColor: '#2c3e50', excelSize: 'XL(검정)',
        avg_monthly: 126.7, monthly: { '2026-03': 105, '2026-04': 167, '2026-05': 92 }, default_stock: 63,
      },
    ],
  },
  {
    id: 'stocking',
    name: '스타킹',
    type: 'manual',
    color: '#8e44ad',
    excelName: '스타킹',
    hasExcelUpload: false,
    sizes: [
      { id: 'M',  label: 'M',  chipColor: '#8e44ad', excelSize: 'M',  default_stock: 0 },
      { id: 'XL', label: 'XL', chipColor: '#8e44ad', excelSize: 'XL', default_stock: 0 },
    ],
  },
];

export const PRODUCT_MAP         = Object.fromEntries(PRODUCTS.map(p => [p.id, p]));
export const CALCULATED_PRODUCTS = PRODUCTS.filter(p => p.type === 'calculated');
export const MANUAL_PRODUCTS     = PRODUCTS.filter(p => p.type === 'manual');
