export const SIZE_COLORS = {
  M1: '#e74c3c',
  M2: '#e6b800',
  L:  '#e67e22',
  XL: '#2c3e50',
};

export const SIZE_LABELS = {
  M1: 'M1 빨강',
  M2: 'M2 노랑',
  L:  'L 주황',
  XL: 'XL 검정',
};

// 창고 - 전산 = 오차 기본값 (UI에서 직접 수정 가능)
export const DEFAULT_OCHA = {
  mediquet: { M1: 22, M2: 35, L: -7, XL: -7 },
  rapband:  { M1:  3, M2:  1, L: -8, XL: 30 },
};

// 최근 3개월(2026-03 ~ 2026-05) 평균 사용량 — 엑셀에서 추출
export const USAGE_DATA = {
  mediquet: [
    {
      size: 'M1',
      avg_monthly: 36.3,
      monthly: { '2026-03': 17, '2026-04': 44, '2026-05': 48 },
      default_stock: 52,
    },
    {
      size: 'M2',
      avg_monthly: 90.7,
      monthly: { '2026-03': 105, '2026-04': 87, '2026-05': 80 },
      default_stock: 73,
    },
    {
      size: 'L',
      avg_monthly: 82.7,
      monthly: { '2026-03': 73, '2026-04': 93, '2026-05': 82 },
      default_stock: 21,
    },
    {
      size: 'XL',
      avg_monthly: 67.3,
      monthly: { '2026-03': 49, '2026-04': 75, '2026-05': 78 },
      default_stock: 96,
    },
  ],
  rapband: [
    {
      size: 'M1',
      avg_monthly: 80.7,
      monthly: { '2026-03': 56, '2026-04': 111, '2026-05': 75 },
      default_stock: 84,
    },
    {
      size: 'M2',
      avg_monthly: 98.3,
      monthly: { '2026-03': 87, '2026-04': 123, '2026-05': 55 },
      default_stock: 111,
    },
    {
      size: 'L',
      avg_monthly: 327.7,
      monthly: { '2026-03': 326, '2026-04': 318, '2026-05': 250 },
      default_stock: 110,
    },
    {
      size: 'XL',
      avg_monthly: 126.7,
      monthly: { '2026-03': 105, '2026-04': 167, '2026-05': 92 },
      default_stock: 63,
    },
  ],
};
