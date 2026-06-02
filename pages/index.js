import Head from 'next/head';
import { useState, useCallback, useEffect } from 'react';
import { USAGE_DATA, SIZE_COLORS, SIZE_LABELS, DEFAULT_OCHA } from '../lib/usageData';
import styles from '../styles/Home.module.css';

const PACK_SIZE = 30;
const FREE_SHIP_BOXES = 6;
const LS_KEY = 'bandOrderStocks';
const SIZES = ['M1', 'M2', 'L', 'XL'];
const PRODUCTS = ['mediquet', 'rapband'];

function calcOrder(avgMonthly, currentStock) {
  const safetyStock = Math.ceil(avgMonthly);
  const shortage = safetyStock - currentStock;
  if (shortage <= 0) return { shortage: 0, packs: 0, qty: 0 };
  const packs = Math.ceil(shortage / PACK_SIZE);
  return { shortage, packs, qty: packs * PACK_SIZE };
}

function defaultStocks() {
  const out = {};
  for (const key of PRODUCTS) {
    out[key] = {};
    USAGE_DATA[key].forEach((item) => { out[key][item.size] = item.default_stock; });
  }
  return out;
}

function defaultJeonsan() {
  return {
    mediquet: { M1: null, M2: null, L: null, XL: null },
    rapband:  { M1: null, M2: null, L: null, XL: null },
  };
}

function defaultAdj() {
  return {
    mediquet: { M1: 0, M2: 0, L: 0, XL: 0 },
    rapband:  { M1: 0, M2: 0, L: 0, XL: 0 },
  };
}

function defaultStocking() {
  return { L: 0, XL: 0 };
}

function defaultOcha() {
  return {
    mediquet: { ...DEFAULT_OCHA.mediquet },
    rapband:  { ...DEFAULT_OCHA.rapband  },
  };
}

async function parseExcelStocks(file) {
  const XLSX = await import('xlsx');
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const result = {};
  for (const sheetName of workbook.SheetNames) {
    const up = sheetName.toUpperCase();
    for (const size of SIZES) {
      if (up.includes(`(${size})`)) {
        const cell = workbook.Sheets[sheetName]['D3'];
        if (cell != null) result[size] = Number(cell.v);
        break;
      }
    }
  }
  return result;
}

function SizeChip({ size, small }) {
  return (
    <span
      className={small ? styles.sizeChipSm : styles.sizeChip}
      style={{ background: SIZE_COLORS[size] }}
    >
      {SIZE_LABELS[size]}
    </span>
  );
}

// ── Upload Section ─────────────────────────────────────────
function UploadSection({ onUpload }) {
  const [medFile, setMedFile] = useState(null);
  const [rapFile, setRapFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState('');

  const handleApply = async () => {
    if (!medFile && !rapFile) { setMsg('파일을 먼저 선택해주세요.'); return; }
    setLoading(true); setMsg('');
    try {
      const results = {};
      if (medFile) results.mediquet = await parseExcelStocks(medFile);
      if (rapFile) results.rapband  = await parseExcelStocks(rapFile);
      onUpload(results);
      setMsg('✅ 재고 자동 업데이트 완료 (전산 + 오차 적용)');
    } catch (e) {
      setMsg('❌ 파일 읽기 오류: ' + e.message);
    }
    setLoading(false);
  };

  return (
    <div className={styles.uploadBox}>
      <h2 className={styles.uploadTitle}>엑셀 파일 업로드</h2>
      <p className={styles.uploadDesc}>업로드하면 전산재고 + 오차값으로 창고재고를 자동 계산합니다.</p>
      <div className={styles.uploadRow}>
        {[
          { key: 'med', label: 'Mediquet', color: '#e74c3c', file: medFile, set: setMedFile },
          { key: 'rap', label: 'Rapband',  color: '#2980b9', file: rapFile, set: setRapFile },
        ].map(({ key, label, color, file, set }) => (
          <label key={key} className={styles.fileLabel}>
            <span className={styles.prodTag} style={{ background: color }}>{label}</span>
            <input type="file" accept=".xlsx,.xls" className={styles.fileInput}
              onChange={(e) => set(e.target.files[0] || null)} />
            <span className={styles.fileName}>{file ? file.name : '파일 선택'}</span>
          </label>
        ))}
      </div>
      <button onClick={handleApply} disabled={loading} className={styles.applyBtn}>
        {loading ? '처리 중...' : '재고 자동 계산 적용'}
      </button>
      {msg && <p className={styles.uploadMsg}>{msg}</p>}
    </div>
  );
}

// ── Ocha (오차) Settings Panel ─────────────────────────────
function OchaPanel({ ocha, onChange, onReset }) {
  const [open, setOpen] = useState(false);

  return (
    <div className={styles.ochaBox}>
      <button className={styles.ochaToggle} onClick={() => setOpen((v) => !v)}>
        <span>⚙️ 오차 설정</span>
        <span className={styles.ochaToggleHint}>
          {open ? '접기 ▲' : '창고재고 − 전산재고 값 편집 ▼'}
        </span>
      </button>

      {open && (
        <div className={styles.ochaBody}>
          <p className={styles.ochaDesc}>
            오차 = 창고 실재고 − 전산 재고량. 엑셀 업로드 시 <strong>창고재고 = 전산 + 오차</strong>로 자동 계산됩니다.
          </p>
          <div className={styles.ochaGrid}>
            {PRODUCTS.map((product) => (
              <div key={product} className={styles.ochaProduct}>
                <div className={styles.ochaProductLabel}
                  style={{ borderColor: product === 'mediquet' ? '#e74c3c' : '#2980b9' }}>
                  {product === 'mediquet' ? 'Mediquet' : 'Rapband'}
                </div>
                <div className={styles.ochaRow}>
                  {SIZES.map((size) => (
                    <div key={size} className={styles.ochaItem}>
                      <SizeChip size={size} small />
                      <input
                        type="number"
                        value={ocha[product][size]}
                        onChange={(e) => onChange(product, size, Number(e.target.value))}
                        className={styles.ochaInput}
                      />
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <button className={styles.ochaResetBtn} onClick={onReset}>
            기본값으로 초기화
          </button>
        </div>
      )}
    </div>
  );
}

// ── Stocking Section ──────────────────────────────────────
const STOCKING_SIZES = ['L', 'XL'];

function StockingSection({ stocking, onChange }) {
  const hasOrder = STOCKING_SIZES.some((s) => stocking[s] > 0);
  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: '#8e44ad' }}>
        스타킹
      </h2>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th style={{ width: '120px' }}>사이즈</th>
              <th>발주수량 <span className={styles.small}>(수기 입력)</span></th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {STOCKING_SIZES.map((size) => (
              <tr key={size} className={stocking[size] > 0 ? styles.needOrder : ''}>
                <td>
                  <span className={styles.stockingChip}>{size}</span>
                </td>
                <td>
                  <input
                    type="number"
                    min="0"
                    value={stocking[size]}
                    onChange={(e) => onChange(size, Number(e.target.value))}
                    className={styles.stockInput}
                  />
                </td>
                <td className={`${styles.num} ${stocking[size] > 0 ? styles.bold : ''}`}>
                  {stocking[size] > 0 ? `${stocking[size]}개` : '—'}
                </td>
              </tr>
            ))}
          </tbody>
          {hasOrder && (
            <tfoot>
              <tr className={styles.totalRow}>
                <td colSpan={2}>총 발주량</td>
                <td className={styles.num}>
                  {STOCKING_SIZES.reduce((s, sz) => s + stocking[sz], 0)}개
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ── Product Table ──────────────────────────────────────────
function ProductTable({ title, color, items, stocks, jeonsan, ocha, adjustments, onStockChange, onAdjust, onResetAdj }) {
  const rows = items.map((item) => {
    const stock = stocks[item.size] ?? item.default_stock;
    const { shortage, packs } = calcOrder(item.avg_monthly, stock);
    const adjDelta  = adjustments[item.size] || 0;
    const finalPacks = Math.max(0, packs + adjDelta);
    const finalQty   = finalPacks * PACK_SIZE;
    const jsVal      = jeonsan[item.size];
    const ochaVal    = ocha[item.size] ?? 0;
    return { ...item, stock, shortage, basePacks: packs, finalPacks, finalQty, jeonsan: jsVal, ocha: ochaVal };
  });

  const totalPacks = rows.reduce((s, r) => s + r.finalPacks, 0);
  const totalBoxes = Math.floor(totalPacks / 2);
  const totalQty   = rows.reduce((s, r) => s + r.finalQty, 0);
  const isOdd      = totalPacks > 0 && totalPacks % 2 !== 0;
  const hasAdj     = SIZES.some((s) => (adjustments[s] || 0) !== 0);

  const orderedRows   = rows.filter((r) => r.finalPacks > 0);
  const biggestRow    = orderedRows.length
    ? orderedRows.reduce((a, b) => b.finalPacks > a.finalPacks ? b : a)
    : null;
  const borderlineRow = orderedRows.length
    ? orderedRows.reduce((a, b) => (a.shortage || 0) <= (b.shortage || 0) ? a : b)
    : null;

  const displayTitle = title === 'mediquet' ? 'Mediquet' : 'Rapband';

  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: color }}>
        {displayTitle}
      </h2>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>사이즈</th>
              <th>창고재고</th>
              <th><span className={styles.thSub}>전산 / 오차</span></th>
              <th>월평균<br />사용량</th>
              <th>적정재고<br />(1배수)</th>
              <th>월별 사용량<br /><span className={styles.small}>3월 / 4월 / 5월</span></th>
              <th>부족량</th>
              <th>발주수량<br /><span className={styles.small}>(30개 단위)</span></th>
              <th>팩수</th>
              <th>박스</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const adjDelta   = adjustments[row.size] || 0;
              const isAdjusted = adjDelta !== 0;
              return (
                <tr key={row.size} className={row.finalPacks > 0 ? styles.needOrder : ''}>
                  <td><SizeChip size={row.size} /></td>
                  <td>
                    <input
                      type="number" min="0" value={row.stock}
                      onChange={(e) => onStockChange(row.size, Number(e.target.value))}
                      className={styles.stockInput}
                    />
                  </td>
                  <td className={styles.jeonsanCell}>
                    {row.jeonsan != null ? (
                      <>
                        <span className={styles.jeonsanVal}>{row.jeonsan}</span>
                        <span className={row.ocha >= 0 ? styles.deltaPos : styles.deltaNeg}>
                          {row.ocha >= 0 ? `+${row.ocha}` : row.ocha}
                        </span>
                      </>
                    ) : <span className={styles.noData}>—</span>}
                  </td>
                  <td className={styles.num}>{row.avg_monthly}</td>
                  <td className={styles.num}>{Math.ceil(row.avg_monthly)}</td>
                  <td className={styles.monthCell}>
                    {row.monthly['2026-03']} / {row.monthly['2026-04']} / {row.monthly['2026-05']}
                  </td>
                  <td className={`${styles.num} ${row.shortage > 0 ? styles.red : styles.greenTxt}`}>
                    {row.shortage > 0 ? `+${row.shortage}` : '충분'}
                  </td>
                  <td className={`${styles.num} ${row.finalQty > 0 ? styles.bold : ''}`}>
                    {row.finalQty > 0 ? `${row.finalQty}개` : '—'}
                    {isAdjusted && (
                      <span className={adjDelta > 0 ? styles.adjTagPos : styles.adjTagNeg}>
                        {adjDelta > 0 ? `+${adjDelta}팩` : `${adjDelta}팩`}
                      </span>
                    )}
                  </td>
                  <td className={`${styles.num} ${row.finalPacks > 0 ? styles.bold : ''}`}>
                    {row.finalPacks > 0 ? `${row.finalPacks}팩` : '—'}
                  </td>
                  <td className={styles.num}>
                    {row.finalPacks > 0
                      ? <span className={row.finalPacks % 2 !== 0 ? styles.halfBox : ''}>
                          {row.finalPacks % 2 === 0
                            ? `${row.finalPacks / 2}박스`
                            : `${Math.floor(row.finalPacks / 2)}+½`}
                        </span>
                      : '—'}
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className={styles.totalRow}>
              <td colSpan={8}>소계</td>
              <td className={styles.num}>{totalPacks > 0 ? `${totalPacks}팩` : '—'}</td>
              <td className={styles.num}>{totalBoxes > 0 ? `${totalBoxes}박스` : '—'}</td>
            </tr>
            <tr className={styles.boxRow}>
              <td colSpan={9}>{displayTitle} 총 박스 수</td>
              <td className={styles.num}>
                <span
                  className={`${styles.boxBadge} ${isOdd ? styles.boxBadgeOdd : ''}`}
                  style={isOdd ? undefined : { background: color }}
                >
                  {isOdd ? `${totalBoxes}+½` : `${totalBoxes}박스`}
                </span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* 홀수 팩 조정 패널 */}
      {isOdd && (
        <div className={styles.oddPanel}>
          <p className={styles.oddTitle}>
            ⚠️ 총 <strong>{totalPacks}팩</strong> (홀수) — 완성 박스 구성 불가. 조정 방법을 선택하세요:
          </p>
          <div className={styles.oddBtns}>
            {biggestRow && (
              <button className={styles.oddBtnReduce} onClick={() => onAdjust(biggestRow.size, -1)}>
                <span className={styles.oddBtnLabel}>줄이기</span>
                <SizeChip size={biggestRow.size} small />
                <span className={styles.oddBtnDetail}>
                  1팩↓ → {totalPacks - 1}팩 = <strong>{(totalPacks - 1) / 2}박스</strong>
                </span>
              </button>
            )}
            {borderlineRow && (
              <button className={styles.oddBtnAdd} onClick={() => onAdjust(borderlineRow.size, 1)}>
                <span className={styles.oddBtnLabel}>늘리기</span>
                <SizeChip size={borderlineRow.size} small />
                <span className={styles.oddBtnDetail}>
                  1팩↑ → {totalPacks + 1}팩 = <strong>{(totalPacks + 1) / 2}박스</strong>
                </span>
              </button>
            )}
          </div>
        </div>
      )}

      {/* 조정 적용 중 표시 */}
      {!isOdd && hasAdj && (
        <div className={styles.adjApplied}>
          <span>조정 적용됨:</span>
          {SIZES.filter((s) => (adjustments[s] || 0) !== 0).map((s) => (
            <span key={s} className={styles.adjAppliedTag}>
              {SIZE_LABELS[s]} {adjustments[s] > 0 ? `+${adjustments[s]}팩` : `${adjustments[s]}팩`}
            </span>
          ))}
          <button className={styles.adjCancelBtn} onClick={onResetAdj}>취소</button>
        </div>
      )}
    </div>
  );
}

// ── Main Page ──────────────────────────────────────────────
export default function Home() {
  const [stocks,      setStocks]      = useState(defaultStocks);
  const [jeonsan,     setJeonsan]     = useState(defaultJeonsan);
  const [adjustments, setAdjustments] = useState(defaultAdj);
  const [ocha,        setOcha]        = useState(defaultOcha);
  const [stocking,    setStocking]    = useState(defaultStocking);
  const [hydrated,    setHydrated]    = useState(false);
  const [downloading, setDownloading] = useState(false);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_KEY);
      if (saved) {
        const p = JSON.parse(saved);
        if (p.stocks)      setStocks(p.stocks);
        if (p.jeonsan)     setJeonsan(p.jeonsan);
        if (p.adjustments) setAdjustments(p.adjustments);
        if (p.ocha)        setOcha(p.ocha);
        if (p.stocking)    setStocking(p.stocking);
      }
    } catch {}
    setHydrated(true);
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    localStorage.setItem(LS_KEY, JSON.stringify({ stocks, jeonsan, adjustments, ocha, stocking }));
  }, [stocks, jeonsan, adjustments, ocha, hydrated]);

  const resetProductAdj = useCallback((product) => {
    setAdjustments((prev) => ({ ...prev, [product]: { M1: 0, M2: 0, L: 0, XL: 0 } }));
  }, []);

  const handleStockChange = useCallback((product, size, val) => {
    setStocks((prev) => ({ ...prev, [product]: { ...prev[product], [size]: val } }));
    resetProductAdj(product);
  }, [resetProductAdj]);

  const handleAdjust = useCallback((product, size, delta) => {
    setAdjustments((prev) => ({
      ...prev,
      [product]: { ...prev[product], [size]: (prev[product][size] || 0) + delta },
    }));
  }, []);

  const handleOchaChange = useCallback((product, size, val) => {
    setOcha((prev) => ({ ...prev, [product]: { ...prev[product], [size]: val } }));
  }, []);

  const handleOchaReset = useCallback(() => {
    setOcha(defaultOcha());
  }, []);

  const handleUpload = useCallback((results) => {
    setStocks((prev) => {
      const next = { ...prev };
      for (const [product, jsMap] of Object.entries(results)) {
        next[product] = { ...next[product] };
        for (const [size, jsVal] of Object.entries(jsMap)) {
          next[product][size] = jsVal + (ocha[product]?.[size] ?? 0);
        }
      }
      return next;
    });
    setJeonsan((prev) => {
      const next = { ...prev };
      for (const [product, jsMap] of Object.entries(results)) {
        next[product] = { ...next[product], ...jsMap };
      }
      return next;
    });
    setAdjustments(defaultAdj());
  }, [ocha]);

  const handleDownload = async () => {
    setDownloading(true);
    try {
      const orders = [];
      for (const product of PRODUCTS) {
        for (const item of USAGE_DATA[product]) {
          const stock = stocks[product][item.size] ?? item.default_stock;
          const { packs: base } = calcOrder(item.avg_monthly, stock);
          const adj = adjustments[product][item.size] || 0;
          const finalPacks = Math.max(0, base + adj);
          const qty = finalPacks * PACK_SIZE;
          if (qty > 0) orders.push({ product, size: item.size, qty });
        }
      }

      // 스타킹
      STOCKING_SIZES.forEach((size) => {
        if (stocking[size] > 0) orders.push({ product: 'stocking', size, qty: stocking[size] });
      });

      if (orders.length === 0) {
        alert('발주할 항목이 없습니다.');
        setDownloading(false);
        return;
      }

      const res = await fetch('/api/generate-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orders }),
      });

      if (!res.ok) throw new Error('서버 오류 ' + res.status);

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const today = new Date();
      const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
      a.href = url;
      a.download = `발주서_${ds}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert('다운로드 실패: ' + e.message);
    }
    setDownloading(false);
  };

  const handleStockingChange = useCallback((size, val) => {
    setStocking((prev) => ({ ...prev, [size]: val }));
  }, []);

  const handleReset = () => {
    setStocks(defaultStocks());
    setJeonsan(defaultJeonsan());
    setAdjustments(defaultAdj());
    setOcha(defaultOcha());
    setStocking(defaultStocking());
  };

  const calcTotalBoxes = (product) => {
    const packs = USAGE_DATA[product].reduce((s, item) => {
      const stock = stocks[product][item.size] ?? item.default_stock;
      const { packs: base } = calcOrder(item.avg_monthly, stock);
      const adj = adjustments[product][item.size] || 0;
      return s + Math.max(0, base + adj);
    }, 0);
    return Math.floor(packs / 2);
  };

  const medBoxes   = calcTotalBoxes('mediquet');
  const rapBoxes   = calcTotalBoxes('rapband');
  const grandTotal = medBoxes + rapBoxes;
  const meetsMin   = grandTotal >= FREE_SHIP_BOXES;

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

          <OchaPanel
            ocha={ocha}
            onChange={handleOchaChange}
            onReset={handleOchaReset}
          />

          <StockingSection stocking={stocking} onChange={handleStockingChange} />

          {PRODUCTS.map((product) => (
            <ProductTable
              key={product}
              title={product}
              color={product === 'mediquet' ? '#e74c3c' : '#2980b9'}
              items={USAGE_DATA[product]}
              stocks={stocks[product]}
              jeonsan={jeonsan[product]}
              ocha={ocha[product]}
              adjustments={adjustments[product]}
              onStockChange={(size, val) => handleStockChange(product, size, val)}
              onAdjust={(size, delta) => handleAdjust(product, size, delta)}
              onResetAdj={() => resetProductAdj(product)}
            />
          ))}

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
            {meetsMin
              ? <p className={styles.summaryMsg}>✅ 무료배송 조건 충족 (6박스 이상)</p>
              : <p className={styles.summaryMsg}>
                  ⚠️ 현재 {grandTotal}박스 — 무료배송까지 <strong>{FREE_SHIP_BOXES - grandTotal}박스</strong> 부족
                </p>
            }
            <button
              onClick={handleDownload}
              disabled={downloading || grandTotal === 0}
              className={styles.downloadBtn}
            >
              {downloading ? '생성 중...' : '📄 발주서 엑셀 다운로드'}
            </button>
          </div>

          <div className={styles.infoBox}>
            <strong>계산 기준</strong>
            <ul>
              <li>적정재고 = 최근 3개월 월평균 사용량 × 1배수</li>
              <li>발주수량 = (적정재고 − 창고재고) 30개 단위 올림</li>
              <li>1팩 = 30개 / 1박스 = 60개 (2팩)</li>
              <li>팩수가 홀수면 테이블 하단에서 조정 방법을 선택할 수 있습니다</li>
              <li>무료배송 기준: 합계 6박스 이상</li>
              <li>오차 = 창고 실재고 − 전산 재고 (상단 오차 설정에서 수정 가능)</li>
            </ul>
          </div>
        </main>
      </div>
    </>
  );
}
