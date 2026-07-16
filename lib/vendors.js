import { PRODUCTS } from './products';

export const TEMPLATES = {
  repmedicare: {
    cells: { date: 'D4', requester: 'M4', recipient: 'D6' },
    itemTable: {
      startRow: 10, endRow: 22,
      cols: { no: 'A', name: 'D', spec: 'F', qty: 'H' },
      groupGap: true, overflow: 'clone-sheet',
    },
    sheetName: '렙메디케어',
  },
  medyssey: {
    cells: { date: 'C3', requester: 'L7', note: 'D30' },
    itemTable: {
      startRow: 10, endRow: 25,
      cols: { no: 'A', name: 'B', spec: 'D', qty: 'E', note: 'N' },
      groupGap: false, overflow: 'clone-sheet',
    },
    sheetName: '22.11.14서울대',
  },
};

export const VENDORS = [
  {
    id: 'repmedicare',
    name: '렙메디케어',
    templateId: 'repmedicare',
    color: '#1a1a2e',
    products: PRODUCTS,
  },
  {
    id: 'spinotech',
    name: '스파이노텍',
    templateId: 'repmedicare',
    color: '#1a6e8a',
    group: 'supplies',
    products: [
      {
        id: 'manan_needle', name: '마난니들', type: 'manual', color: '#1a6e8a',
        excelName: '마난니들', hasExcelUpload: false,
        sizes: [{ id: 'single', label: '단일', chipColor: '#1a6e8a', excelSize: '', default_stock: 0 }],
      },
    ],
  },
  {
    id: 'tdm',
    name: 'TDM',
    templateId: 'repmedicare',
    color: '#16a085',
    group: 'supplies',
    products: [
      {
        id: 'surgical_gown', name: '수술복', type: 'manual', color: '#16a085',
        excelName: '수술복', hasExcelUpload: false,
        sizes: [{ id: 'single', label: '단일', chipColor: '#16a085', excelSize: '', default_stock: 0 }],
      },
      {
        id: 'cable', name: '케이블', type: 'manual', color: '#27ae60',
        excelName: '케이블', hasExcelUpload: false,
        sizes: [{ id: 'single', label: '단일', chipColor: '#27ae60', excelSize: '', default_stock: 0 }],
      },
    ],
  },
  {
    id: 'bioimetec',
    name: '바이오임텍',
    templateId: 'repmedicare',
    color: '#8e44ad',
    group: 'supplies',
    products: [
      {
        id: 'spine_fix', name: '스파인픽스', type: 'manual', color: '#8e44ad',
        excelName: '스파인픽스', hasExcelUpload: false,
        sizes: [{ id: 'single', label: '단일', chipColor: '#8e44ad', excelSize: '', default_stock: 0 }],
      },
      {
        id: 'redura', name: '레듀라', type: 'manual', color: '#9b59b6',
        excelName: '레듀라', hasExcelUpload: false,
        sizes: [{ id: 'single', label: '단일', chipColor: '#9b59b6', excelSize: '', default_stock: 0 }],
      },
    ],
  },
  {
    id: 'bijet',
    name: 'Bi-Jet',
    color: '#0e7c66',
    type: 'usage',
    usageMonths: ['2026-04', '2026-05', '2026-06'],
    products: [
      {
        id: 'bijet_external', name: '석션(외장형)', color: '#0e7c66',
        avg_monthly: 745,
        monthly: { '2026-04': 1155, '2026-05': 205, '2026-06': 875 },
      },
      {
        id: 'bijet_plus_sterile', name: 'Bijet Plus(멸균)', color: '#1a8f76',
        avg_monthly: 431,
        monthly: { '2026-04': 570, '2026-05': 365, '2026-06': 358 },
      },
    ],
  },
  {
    id: 'medyssey',
    name: '메디쎄이',
    templateId: 'medyssey',
    color: '#a8531b',
    type: 'history',
    hasHospitalLayer: true,
    products: [],
  },
];

export const VENDOR_MAP = Object.fromEntries(VENDORS.map(v => [v.id, v]));
