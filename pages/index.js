import Head from 'next/head';
import { useState, useCallback, useEffect, useRef } from 'react';
import { USAGE_DATA, SIZE_COLORS, SIZE_LABELS, GEUMJU_DELTA } from '../lib/usageData';
import styles from '../styles/Home.module.css';

const PACK_SIZE = 30;
const FREE_SHIP_BOXES = 6;
const LS_KEY = 'bandOrderStocks';

function calcOrder(avgMonthly, currentStock) {
  const safetyStock = Math.ceil(avgMonthly);
  const shortage = safetyStock - currentStock;
  if (shortage <= 0) return { shortage: 0, packs: 0, qty: 0 };
  const packs = Math.ceil(shortage / PACK_SIZE);
  return { shortage, packs, qty: packs * PACK_SIZE };
}

function defaultStocks() {
  const out = {};
  for (const key of ['mediquet', 'rapband']) {
    out[key] = {};
    USAGE_DATA[key].forEach((item) => { out[key][item.size] = item.default_stock; });
  }
  return out;
}

function defaultJeonsan() {
  return { mediquet: { M1: null, M2: null, L: null, XL: null }, rapband: { M1: null, M2: null, L: null, XL: null } };
}

// 엑셀 시트에서 현재고량(D3) 읽기
async function parseExcelStocks(file, productKey) {
  const XLSX = await import('xlsx');
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const result = {};
  const sizes = ['M1', 'M2', 'L', 'XL'];

  for (const sheetName of workbook.SheetNames) {
    const upperSheet = sheetName.toUpperCase();
    for (const size of sizes) {
      // 시트명에 사이즈 포함 여부 확인 ("(M1)", "(L)" 형태)
      if (upperSheet.includes(`(${size})`)) {
        const ws = workbook.Sheets[sheetName];
        const cell = ws['D3'];
        if (cell != null) {
          result[size] = Number(cell.v);
        }
        break;
      }
    }
  }
  return result;
}

function SizeChip({ size }) {
  return (
    <span className={styles.sizeChip} style={{ background: SIZE_COLORS[size] }}>
      {SIZE_LABELS[size]}
    </span>
  );
}

function UploadSection({ onUpload }) {
  const medRef = useRef(null);
  const rapRef = useRef(null);
  const [medFile, setMedFile] = useState(null);
  const [rapFile, setRapFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState('');

  const handleApply = async () => {
    if (!medFile && !rapFile) { setMsg('파일을 먼저 선택해주세요.'); return; }
    setLoading(true);
    setMsg('');
    try {
      const results = {};
      if (medFile) results.mediquet = await parseExcelStocks(medFile, 'mediquet');
      if (rapFile) results.rapband  = await parseExcelStocks(rapFile, 'rapband');
      onUpload(results);
      setMsg('✅ 재고 자동 업데이트 완료 (전산 + 금주 조정 적용)');
    } catch (e) {
      setMsg('❌ 파일 읽기 오류: ' + e.message);
    }
    setLoading(false);
  };

  return (
    <div className={styles.uploadBox}>
      <h2 className={styles.uploadTitle}>엑셀 파일 업로드</h2>
      <p className={styles.uploadDesc}>
        업로드하면 전산 재고 + 금주 조정값으로 창고재고를 자동 계산합니다.
      </p>
      <div className={styles.uploadRow}>
        <label className={styles.fileLabel}>
          <span className={styles.prodTag} style={{ background: '#e74c3c' }}>Mediquet</span>
          <input
            ref={medRef}
            type="file"
            accept=".xlsx,.xls"
            className={styles.fileInput}
            onChange={(e) => setMedFile(e.target.files[0] || null)}
          />
          <span className={styles.fileName}>{medFile ? medFile.name : '파일 선택'}</span>
        </label>
        <label className={styles.fileLabel}>
          <span className={styles.prodTag} style={{ background: '#2980b9' }}>Rapband</span>
          <input
            ref={rapRef}
            type="file"
            accept=".xlsx,.xls"
            className={styles.fileInput}
            onChange={(e) => setRapFile(e.target.files[0] || null)}
          />
          <span className={styles.fileName}>{rapFile ? rapFile.name : '파일 선택'}</span>
        </label>
      </div>
      <button onClick={handleApply} disabled={loading} className={styles.applyBtn}>
        {loading ? '처리 중...' : '재고 자동 계산 적용'}
      </button>
      {msg && <p className={styles.uploadMsg}>{msg}</p>}
    </div>
  );
}

function ProductTable({ title, color, items, stocks, jeonsan, onStockChange }) {
  const rows = items.map((item) => {
    const stock = stocks[item.size] ?? item.default_stock;
    const { shortage, packs, qty } = calcOrder(item.avg_monthly, stock);
    const jsVal = jeonsan[item.size];
    const delta = GEUMJU_DELTA[title.toLowerCase()] ?? {};
    return { ...item, stock, shortage, packs, qty, jeonsan: jsVal, delta: delta[item.size] ?? 0 };
  });

  const totalPacks = rows.reduce((s, r) => s + r.packs, 0);
  const totalBoxes = Math.ceil(totalPacks / 2);
  const totalQty   = rows.reduce((s, r) => s + r.qty, 0);

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
              <th>
                <span className={styles.thSub}>전산</span>
                <br />
                <span className={styles.thSub}>금주</span>
              </th>
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
                <td><SizeChip size={row.size} /></td>
                <td>
                  <input
                    type="number"
                    min="0"
                    value={row.stock}
                    onChange={(e) => onStockChange(row.size, Number(e.target.value))}
                    className={styles.stockInput}
                  />
                </td>
                <td className={styles.jeonsanCell}>
                  {row.jeonsan != null ? (
                    <>
                      <span className={styles.jeonsanVal}>{row.jeonsan}</span>
                      <span className={row.delta >= 0 ? styles.deltaPos : styles.deltaNeg}>
                        {row.delta >= 0 ? `+${row.delta}` : row.delta}
                      </span>
                    </>
                  ) : (
                    <span className={styles.noData}>—</span>
                  )}
                </td>
                <td className={styles.num}>{row.avg_monthly}</td>
                <td className={styles.num}>{Math.ceil(row.avg_monthly)}</td>
                <td className={styles.monthCell}>
                  {row.monthly['2026-03']} / {row.monthly['2026-04']} / {row.monthly['2026-05']}
                </td>
                <td className={`${styles.num} ${row.shortage > 0 ? styles.red : styles.greenTxt}`}>
                  {row.shortage > 0 ? `+${row.shortage}` : '충분'}
                </td>
                <td className={`${styles.num} ${row.qty > 0 ? styles.bold : ''}`}>
                  {row.qty > 0 ? `${row.qty}개` : '—'}
                </td>
                <td className={`${styles.num} ${row.packs > 0 ? styles.bold : ''}`}>
                  {row.packs > 0 ? `${row.packs}팩` : '—'}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className={styles.totalRow}>
              <td colSpan={7}>소계</td>
              <td className={styles.num}>{totalQty > 0 ? `${totalQty}개` : '—'}</td>
              <td className={styles.num}>{totalPacks > 0 ? `${totalPacks}팩` : '—'}</td>
            </tr>
            <tr className={styles.boxRow}>
              <td colSpan={8}>{title} 총 박스 수</td>
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
  const [stocks, setStocks] = useState(defaultStocks);
  const [jeonsan, setJeonsan] = useState(defaultJeonsan);
  const [hydrated, setHydrated] = useState(false);

  // localStorage 복원
  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.stocks) setStocks(parsed.stocks);
        if (parsed.jeonsan) setJeonsan(parsed.jeonsan);
      }
    } catch {}
    setHydrated(true);
  }, []);

  // localStorage 저장
  useEffect(() => {
    if (!hydrated) return;
    localStorage.setItem(LS_KEY, JSON.stringify({ stocks, jeonsan }));
  }, [stocks, jeonsan, hydrated]);

  const handleStockChange = useCallback((product, size, val) => {
    setStocks((prev) => ({ ...prev, [product]: { ...prev[product], [size]: val } }));
  }, []);

  const handleUpload = useCallback((results) => {
    setStocks((prev) => {
      const next = { ...prev };
      for (const [product, jeonsanMap] of Object.entries(results)) {
        const delta = GEUMJU_DELTA[product] ?? {};
        next[product] = { ...next[product] };
        for (const [size, jsVal] of Object.entries(jeonsanMap)) {
          next[product][size] = jsVal + (delta[size] ?? 0);
        }
      }
      return next;
    });
    setJeonsan((prev) => {
      const next = { ...prev };
      for (const [product, jeonsanMap] of Object.entries(results)) {
        next[product] = { ...next[product], ...jeonsanMap };
      }
      return next;
    });
  }, []);

  const handleReset = () => {
    setStocks(defaultStocks());
    setJeonsan(defaultJeonsan());
  };

  const calcTotalBoxes = (product) => {
    const packs = USAGE_DATA[product].reduce((s, item) => {
      const stock = stocks[product][item.size] ?? item.default_stock;
      return s + calcOrder(item.avg_monthly, stock).packs;
    }, 0);
    return Math.ceil(packs / 2);
  };

  const medBoxes = calcTotalBoxes('mediquet');
  const rapBoxes = calcTotalBoxes('rapband');
  const grandTotal = medBoxes + rapBoxes;
  const meetsMin = grandTotal >= FREE_SHIP_BOXES;

  if (!hydrated) return null;

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
          <button onClick={handleReset} className={styles.resetBtn}>초기화</button>
        </header>

        <main className={styles.main}>
          <UploadSection onUpload={handleUpload} />

          <ProductTable
            title="mediquet"
            color="#e74c3c"
            items={USAGE_DATA.mediquet}
            stocks={stocks.mediquet}
            jeonsan={jeonsan.mediquet}
            onStockChange={(size, val) => handleStockChange('mediquet', size, val)}
          />
          <ProductTable
            title="rapband"
            color="#2980b9"
            items={USAGE_DATA.rapband}
            stocks={stocks.rapband}
            jeonsan={jeonsan.rapband}
            onStockChange={(size, val) => handleStockChange('rapband', size, val)}
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
                ⚠️ 현재 {grandTotal}박스 — 무료배송까지 <strong>{FREE_SHIP_BOXES - grandTotal}박스</strong> 부족
              </p>
            )}
          </div>

          <div className={styles.infoBox}>
            <strong>계산 기준</strong>
            <ul>
              <li>적정재고 = 최근 3개월 월평균 사용량 × 1배수</li>
              <li>발주수량 = (적정재고 − 창고재고) 30개 단위 올림</li>
              <li>1팩 = 30개 / 1박스 = 60개 (2팩)</li>
              <li>무료배송 기준: 합계 6박스 이상</li>
              <li>금주 조정 = 창고재고 − 전산재고 (변동 시 알려주세요)</li>
              <li>사용량 데이터: 2026-03 ~ 2026-05</li>
            </ul>
          </div>
        </main>
      </div>
    </>
  );
}
