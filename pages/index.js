import Head from 'next/head';
import { useState, useCallback } from 'react';
import { USAGE_DATA } from '../lib/usageData';
import styles from '../styles/Home.module.css';

const PACK_SIZE = 30;
const FREE_SHIP_BOXES = 6;

function calcOrder(avgMonthly, currentStock) {
  const safetyStock = Math.ceil(avgMonthly);
  const shortage = safetyStock - currentStock;
  if (shortage <= 0) return { shortage: 0, packs: 0, qty: 0 };
  const packs = Math.ceil(shortage / PACK_SIZE);
  return { shortage, packs, qty: packs * PACK_SIZE };
}

function ProductTable({ title, color, items, stocks, onStockChange }) {
  const rows = items.map((item) => {
    const stock = stocks[item.size] ?? item.default_stock;
    const { shortage, packs, qty } = calcOrder(item.avg_monthly, stock);
    return { ...item, stock, shortage, packs, qty };
  });

  const totalPacks = rows.reduce((s, r) => s + r.packs, 0);
  const totalBoxes = Math.ceil(totalPacks / 2);
  const totalQty = rows.reduce((s, r) => s + r.qty, 0);

  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: color }}>
        {title}
      </h2>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>사이즈</th>
              <th>창고재고</th>
              <th>월평균<br />사용량</th>
              <th>적정재고<br />(1배수)</th>
              <th>월별 사용량<br /><span className={styles.small}>3월 / 4월 / 5월</span></th>
              <th>부족량</th>
              <th>발주수량<br /><span className={styles.small}>(30개 단위)</span></th>
              <th>팩수</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.size} className={row.packs > 0 ? styles.needOrder : ''}>
                <td className={styles.sizeCell}>{row.label}</td>
                <td>
                  <input
                    type="number"
                    min="0"
                    value={row.stock}
                    onChange={(e) => onStockChange(row.size, Number(e.target.value))}
                    className={styles.stockInput}
                  />
                </td>
                <td className={styles.num}>{row.avg_monthly}</td>
                <td className={styles.num}>{Math.ceil(row.avg_monthly)}</td>
                <td className={styles.monthCell}>
                  {row.monthly['2026-03']} / {row.monthly['2026-04']} / {row.monthly['2026-05']}
                </td>
                <td className={`${styles.num} ${row.shortage > 0 ? styles.red : styles.green}`}>
                  {row.shortage > 0 ? `+${row.shortage}` : '충분'}
                </td>
                <td className={`${styles.num} ${row.qty > 0 ? styles.bold : ''}`}>
                  {row.qty > 0 ? `${row.qty}개` : '-'}
                </td>
                <td className={`${styles.num} ${row.packs > 0 ? styles.bold : ''}`}>
                  {row.packs > 0 ? `${row.packs}팩` : '-'}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className={styles.totalRow}>
              <td colSpan={6}>소계</td>
              <td className={styles.num}>{totalQty > 0 ? `${totalQty}개` : '-'}</td>
              <td className={styles.num}>{totalPacks > 0 ? `${totalPacks}팩` : '-'}</td>
            </tr>
            <tr className={styles.boxRow}>
              <td colSpan={7}>{title} 박스 수</td>
              <td className={styles.num}>
                <span className={styles.boxBadge} style={{ background: color }}>
                  {totalBoxes}박스
                </span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

export default function Home() {
  const initStocks = (items) =>
    Object.fromEntries(items.map((i) => [i.size, i.default_stock]));

  const [medStocks, setMedStocks] = useState(() => initStocks(USAGE_DATA.mediquet));
  const [rapStocks, setRapStocks] = useState(() => initStocks(USAGE_DATA.rapband));

  const handleMedStock = useCallback((size, val) => {
    setMedStocks((prev) => ({ ...prev, [size]: val }));
  }, []);

  const handleRapStock = useCallback((size, val) => {
    setRapStocks((prev) => ({ ...prev, [size]: val }));
  }, []);

  const calcTotalBoxes = (items, stocks) => {
    const packs = items.reduce((s, item) => {
      const stock = stocks[item.size] ?? item.default_stock;
      return s + calcOrder(item.avg_monthly, stock).packs;
    }, 0);
    return Math.ceil(packs / 2);
  };

  const medBoxes = calcTotalBoxes(USAGE_DATA.mediquet, medStocks);
  const rapBoxes = calcTotalBoxes(USAGE_DATA.rapband, rapStocks);
  const grandTotal = medBoxes + rapBoxes;
  const meetsMin = grandTotal >= FREE_SHIP_BOXES;
  const shortfall = FREE_SHIP_BOXES - grandTotal;

  const handleReset = () => {
    setMedStocks(initStocks(USAGE_DATA.mediquet));
    setRapStocks(initStocks(USAGE_DATA.rapband));
  };

  return (
    <>
      <Head>
        <title>랩밴드 / 메디켓 발주 관리</title>
        <meta name="description" content="랩밴드 메디켓 발주수량 계산기" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="icon" href="/favicon.ico" />
      </Head>
      <div className={styles.page}>
        <header className={styles.header}>
          <h1 className={styles.title}>랩밴드 / 메디켓 발주 관리</h1>
          <p className={styles.subtitle}>최근 3개월 월평균 사용량 기준 · 적정재고 1배수 · 30개 단위 발주</p>
          <button onClick={handleReset} className={styles.resetBtn}>재고 초기화</button>
        </header>

        <main className={styles.main}>
          <ProductTable
            title="Mediquet"
            color="#e74c3c"
            items={USAGE_DATA.mediquet}
            stocks={medStocks}
            onStockChange={handleMedStock}
          />
          <ProductTable
            title="Rapband"
            color="#2980b9"
            items={USAGE_DATA.rapband}
            stocks={rapStocks}
            onStockChange={handleRapStock}
          />

          <div className={`${styles.summary} ${meetsMin ? styles.summaryOk : styles.summaryWarn}`}>
            <div className={styles.summaryRow}>
              <div className={styles.summaryBlock}>
                <span className={styles.summaryLabel}>Mediquet</span>
                <span className={styles.summaryVal}>{medBoxes}박스</span>
              </div>
              <span className={styles.summaryPlus}>+</span>
              <div className={styles.summaryBlock}>
                <span className={styles.summaryLabel}>Rapband</span>
                <span className={styles.summaryVal}>{rapBoxes}박스</span>
              </div>
              <span className={styles.summaryPlus}>=</span>
              <div className={styles.summaryBlock}>
                <span className={styles.summaryLabel}>총 발주</span>
                <span className={`${styles.summaryVal} ${styles.summaryTotal}`}>{grandTotal}박스</span>
              </div>
            </div>
            {meetsMin ? (
              <p className={styles.summaryMsg}>✅ 무료배송 조건 충족 (6박스 이상)</p>
            ) : (
              <p className={styles.summaryMsg}>
                ⚠️ 현재 {grandTotal}박스 — 무료배송까지 <strong>{shortfall}박스</strong> 부족
              </p>
            )}
          </div>

          <div className={styles.infoBox}>
            <strong>계산 기준</strong>
            <ul>
              <li>적정재고 = 최근 3개월 월평균 사용량 × 1배수</li>
              <li>발주수량 = (적정재고 − 창고재고)를 30개 단위 올림</li>
              <li>1팩 = 30개 / 1박스 = 60개 (2팩)</li>
              <li>무료배송 기준: 합계 6박스 이상</li>
              <li>사용량 데이터: 2026-03 ~ 2026-05 (엑셀 기준)</li>
            </ul>
          </div>
        </main>
      </div>
    </>
  );
}
