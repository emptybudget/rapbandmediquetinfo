import Head from 'next/head';
import { useState, useCallback, useEffect, Fragment } from 'react';
import {
  PRODUCTS, CALCULATED_PRODUCTS, MANUAL_PRODUCTS,
  PACK_SIZE, FREE_SHIP_BOXES,
} from '../lib/products';
import styles from '../styles/Home.module.css';

// ── Supply Order Tab ───────────────────────────────────────
const EMPTY_ITEM = () => ({ name: '', spec: '', qty: 1 });
const LS_SUPPLY = 'supplyOrderDraft';

function SupplyOrderTab() {
  const [requester, setRequester] = useState('');
  const [recipient, setRecipient] = useState('');
  const [items, setItems] = useState([EMPTY_ITEM()]);
  const [hydrated, setHydrated] = useState(false);
  const [downloading, setDownloading] = useState(false);

  // Load from localStorage on mount
  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_SUPPLY);
      if (saved) {
        const p = JSON.parse(saved);
        if (p.requester !== undefined) setRequester(p.requester);
        if (p.recipient !== undefined) setRecipient(p.recipient);
        if (p.items?.length)           setItems(p.items);
      }
    } catch {}
    setHydrated(true);
  }, []);

  // Auto-save on every change
  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(LS_SUPPLY, JSON.stringify({ requester, recipient, items })); } catch {}
  }, [requester, recipient, items, hydrated]);

  const addItem = () => setItems(prev => [...prev, EMPTY_ITEM()]);
  const removeItem = (i) => setItems(prev => prev.filter((_, idx) => idx !== i));
  const updateItem = (i, field, val) =>
    setItems(prev => prev.map((item, idx) => idx === i ? { ...item, [field]: val } : item));

  const handleDownload = async () => {
    const validItems = items.filter(it => it.name.trim());
    if (!validItems.length) { alert('품명을 입력해주세요.'); return; }
    setDownloading(true);
    try {
      const res = await fetch('/api/generate-supply-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requester, recipient, items: validItems }),
      });
      if (!res.ok) throw new Error('서버 오류 ' + res.status);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const today = new Date();
      const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
      a.href = url;
      a.download = `${validItems[0].name}_발주서_${ds}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert('다운로드 실패: ' + e.message);
    }
    setDownloading(false);
  };

  return (
    <div className={styles.supplyWrap}>
      {/* Header info */}
      <div className={styles.section}>
        <h2 className={styles.sectionTitle} style={{ borderLeftColor: '#2c7be5' }}>발주 정보</h2>
        <div className={styles.supplyHeaderGrid}>
          <div className={styles.supplyField}>
            <label className={styles.supplyLabel}>출고의뢰인</label>
            <input
              type="text"
              value={requester}
              onChange={e => setRequester(e.target.value)}
              placeholder="이름 입력"
              className={styles.supplyInput}
            />
          </div>
          <div className={styles.supplyField}>
            <label className={styles.supplyLabel}>수주처 명</label>
            <input
              type="text"
              value={recipient}
              onChange={e => setRecipient(e.target.value)}
              placeholder="수주처 입력"
              className={styles.supplyInput}
            />
          </div>
        </div>
      </div>

      {/* Items table */}
      <div className={styles.section}>
        <h2 className={styles.sectionTitle} style={{ borderLeftColor: '#2c7be5' }}>발주 품목</h2>
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th style={{ width: 36 }}>No.</th>
                <th>품명</th>
                <th style={{ width: 160 }}>규격</th>
                <th style={{ width: 90 }}>수량</th>
                <th style={{ width: 44 }}></th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, i) => (
                <tr key={i}>
                  <td className={styles.num}>{i + 1}</td>
                  <td>
                    <input
                      type="text"
                      value={item.name}
                      onChange={e => updateItem(i, 'name', e.target.value)}
                      placeholder="품명"
                      className={styles.supplyInputWide}
                    />
                  </td>
                  <td>
                    <input
                      type="text"
                      value={item.spec}
                      onChange={e => updateItem(i, 'spec', e.target.value)}
                      placeholder="규격"
                      className={styles.supplyInputMid}
                    />
                  </td>
                  <td>
                    <input
                      type="number"
                      min="0"
                      value={item.qty}
                      onChange={e => updateItem(i, 'qty', e.target.value)}
                      className={styles.stockInput}
                    />
                  </td>
                  <td>
                    {items.length > 1 && (
                      <button
                        onClick={() => removeItem(i)}
                        className={styles.supplyRemoveBtn}
                        title="삭제"
                      >✕</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className={styles.supplyFooter}>
          {items.length < 13 && (
            <button onClick={addItem} className={styles.supplyAddBtn}>+ 품목 추가</button>
          )}
          <span className={styles.supplyCount}>{items.length} / 13행</span>
        </div>
      </div>

      {/* Download */}
      <button
        onClick={handleDownload}
        disabled={downloading}
        className={styles.downloadBtn}
        style={{ marginTop: 0 }}
      >
        {downloading ? '생성 중...' : '📄 발주서 엑셀 다운로드'}
      </button>
    </div>
  );
}

const LS_KEY = 'bandOrderStocks';

function calcOrder(avgMonthly, currentStock) {
  const safetyStock = Math.ceil(avgMonthly);
  const shortage = safetyStock - currentStock;
  if (shortage <= 0) return { shortage: 0, packs: 0, qty: 0 };
  const packs = Math.ceil(shortage / PACK_SIZE);
  return { shortage, packs, qty: packs * PACK_SIZE };
}

function defaultStocks() {
  return Object.fromEntries(
    CALCULATED_PRODUCTS.map(p => [p.id, Object.fromEntries(p.sizes.map(s => [s.id, s.default_stock]))])
  );
}

function defaultJeonsan() {
  return Object.fromEntries(
    CALCULATED_PRODUCTS.map(p => [p.id, Object.fromEntries(p.sizes.map(s => [s.id, null]))])
  );
}

function defaultAdj() {
  return Object.fromEntries(
    CALCULATED_PRODUCTS.map(p => [p.id, Object.fromEntries(p.sizes.map(s => [s.id, 0]))])
  );
}

function defaultOcha() {
  return Object.fromEntries(
    CALCULATED_PRODUCTS.map(p => [p.id, { ...p.defaultOcha }])
  );
}

function defaultManualOrders() {
  return Object.fromEntries(
    MANUAL_PRODUCTS.map(p => [p.id, Object.fromEntries(p.sizes.map(s => [s.id, s.default_stock]))])
  );
}

async function parseExcelStocks(file, product) {
  const XLSX = await import('xlsx');
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
  const result = {};
  for (const sheetName of workbook.SheetNames) {
    const up = sheetName.toUpperCase();
    for (const sz of product.sizes) {
      if (up.includes(`(${sz.id})`)) {
        const cell = workbook.Sheets[sheetName]['D3'];
        if (cell != null) result[sz.id] = Number(cell.v);
        break;
      }
    }
  }
  return result;
}

function SizeChip({ sizeObj, small }) {
  return (
    <span
      className={small ? styles.sizeChipSm : styles.sizeChip}
      style={{ background: sizeObj.chipColor }}
    >
      {sizeObj.label}
    </span>
  );
}

// ── Upload Section ─────────────────────────────────────────
const UPLOAD_PRODUCTS = PRODUCTS.filter(p => p.hasExcelUpload);

function UploadSection({ onUpload }) {
  const [files, setFiles] = useState(() =>
    Object.fromEntries(UPLOAD_PRODUCTS.map(p => [p.id, null]))
  );
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState('');

  const handleApply = async () => {
    if (!UPLOAD_PRODUCTS.some(p => files[p.id])) { setMsg('파일을 먼저 선택해주세요.'); return; }
    setLoading(true); setMsg('');
    try {
      const results = {};
      for (const prod of UPLOAD_PRODUCTS) {
        if (files[prod.id]) results[prod.id] = await parseExcelStocks(files[prod.id], prod);
      }
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
        {UPLOAD_PRODUCTS.map(prod => (
          <label key={prod.id} className={styles.fileLabel}>
            <span className={styles.prodTag} style={{ background: prod.color }}>{prod.name}</span>
            <input
              type="file" accept=".xlsx,.xls" className={styles.fileInput}
              onChange={e => setFiles(prev => ({ ...prev, [prod.id]: e.target.files[0] || null }))}
            />
            <span className={styles.fileName}>{files[prod.id] ? files[prod.id].name : '파일 선택'}</span>
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
      <button className={styles.ochaToggle} onClick={() => setOpen(v => !v)}>
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
            {CALCULATED_PRODUCTS.map(prod => (
              <div key={prod.id} className={styles.ochaProduct}>
                <div className={styles.ochaProductLabel} style={{ borderColor: prod.color }}>
                  {prod.name}
                </div>
                <div className={styles.ochaRow}>
                  {prod.sizes.map(sz => (
                    <div key={sz.id} className={styles.ochaItem}>
                      <SizeChip sizeObj={sz} small />
                      <input
                        type="number"
                        value={ocha[prod.id]?.[sz.id] ?? 0}
                        onChange={e => onChange(prod.id, sz.id, Number(e.target.value))}
                        className={styles.ochaInput}
                      />
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <button className={styles.ochaResetBtn} onClick={onReset}>기본값으로 초기화</button>
        </div>
      )}
    </div>
  );
}

// ── Manual Product Section (수기 발주 — e.g. 스타킹) ─────
function ManualProductSection({ product, orders, onChange }) {
  const hasOrder = product.sizes.some(sz => (orders[sz.id] || 0) > 0);
  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: product.color }}>
        {product.name}
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
            {product.sizes.map(sz => {
              const qty = orders[sz.id] || 0;
              return (
                <tr key={sz.id} className={qty > 0 ? styles.needOrder : ''}>
                  <td><SizeChip sizeObj={sz} /></td>
                  <td>
                    <input
                      type="number" min="0" value={qty}
                      onChange={e => onChange(product.id, sz.id, Number(e.target.value))}
                      className={styles.stockInput}
                    />
                  </td>
                  <td className={`${styles.num} ${qty > 0 ? styles.bold : ''}`}>
                    {qty > 0 ? `${qty}개` : '—'}
                  </td>
                </tr>
              );
            })}
          </tbody>
          {hasOrder && (
            <tfoot>
              <tr className={styles.totalRow}>
                <td colSpan={2}>총 발주량</td>
                <td className={styles.num}>
                  {product.sizes.reduce((s, sz) => s + (orders[sz.id] || 0), 0)}개
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ── Product Table (calculated products) ───────────────────
function ProductTable({ product, stocks, jeonsan, ocha, adjustments, onStockChange, onAdjust, onResetAdj }) {
  const monthKeys = product.sizes[0]?.monthly ? Object.keys(product.sizes[0].monthly) : [];

  const rows = product.sizes.map(sz => {
    const stock      = stocks[sz.id] ?? sz.default_stock;
    const { shortage, packs } = calcOrder(sz.avg_monthly, stock);
    const adjDelta   = adjustments[sz.id] || 0;
    const finalPacks = Math.max(0, packs + adjDelta);
    const finalQty   = finalPacks * PACK_SIZE;
    const jsVal      = jeonsan[sz.id];
    const ochaVal    = ocha[sz.id] ?? 0;
    return { ...sz, stock, shortage, basePacks: packs, finalPacks, finalQty, jeonsan: jsVal, ocha: ochaVal };
  });

  const totalPacks = rows.reduce((s, r) => s + r.finalPacks, 0);
  const totalBoxes = Math.floor(totalPacks / 2);
  const isOdd      = totalPacks > 0 && totalPacks % 2 !== 0;
  const hasAdj     = product.sizes.some(sz => (adjustments[sz.id] || 0) !== 0);

  const orderedRows   = rows.filter(r => r.finalPacks > 0);
  const biggestRow    = orderedRows.length ? orderedRows.reduce((a, b) => b.finalPacks > a.finalPacks ? b : a) : null;
  const borderlineRow = orderedRows.length ? orderedRows.reduce((a, b) => (a.shortage || 0) <= (b.shortage || 0) ? a : b) : null;

  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: product.color }}>
        {product.name}
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
              <th>
                월별 사용량<br />
                <span className={styles.small}>
                  {monthKeys.map(k => k.split('-')[1].replace(/^0/, '') + '월').join(' / ')}
                </span>
              </th>
              <th>부족량</th>
              <th>발주수량<br /><span className={styles.small}>(30개 단위)</span></th>
              <th>팩수</th>
              <th>박스</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => {
              const adjDelta   = adjustments[row.id] || 0;
              const isAdjusted = adjDelta !== 0;
              return (
                <tr key={row.id} className={row.finalPacks > 0 ? styles.needOrder : ''}>
                  <td><SizeChip sizeObj={row} /></td>
                  <td>
                    <input
                      type="number" min="0" value={row.stock}
                      onChange={e => onStockChange(row.id, Number(e.target.value))}
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
                    {monthKeys.map(k => row.monthly[k]).join(' / ')}
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
              <td colSpan={9}>{product.name} 총 박스 수</td>
              <td className={styles.num}>
                <span
                  className={`${styles.boxBadge} ${isOdd ? styles.boxBadgeOdd : ''}`}
                  style={isOdd ? undefined : { background: product.color }}
                >
                  {isOdd ? `${totalBoxes}+½` : `${totalBoxes}박스`}
                </span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>

      {isOdd && (
        <div className={styles.oddPanel}>
          <p className={styles.oddTitle}>
            ⚠️ 총 <strong>{totalPacks}팩</strong> (홀수) — 완성 박스 구성 불가. 조정 방법을 선택하세요:
          </p>
          <div className={styles.oddBtns}>
            {biggestRow && (
              <button className={styles.oddBtnReduce} onClick={() => onAdjust(biggestRow.id, -1)}>
                <span className={styles.oddBtnLabel}>줄이기</span>
                <SizeChip sizeObj={biggestRow} small />
                <span className={styles.oddBtnDetail}>
                  1팩↓ → {totalPacks - 1}팩 = <strong>{(totalPacks - 1) / 2}박스</strong>
                </span>
              </button>
            )}
            {borderlineRow && (
              <button className={styles.oddBtnAdd} onClick={() => onAdjust(borderlineRow.id, 1)}>
                <span className={styles.oddBtnLabel}>늘리기</span>
                <SizeChip sizeObj={borderlineRow} small />
                <span className={styles.oddBtnDetail}>
                  1팩↑ → {totalPacks + 1}팩 = <strong>{(totalPacks + 1) / 2}박스</strong>
                </span>
              </button>
            )}
          </div>
        </div>
      )}

      {!isOdd && hasAdj && (
        <div className={styles.adjApplied}>
          <span>조정 적용됨:</span>
          {product.sizes.filter(sz => (adjustments[sz.id] || 0) !== 0).map(sz => (
            <span key={sz.id} className={styles.adjAppliedTag}>
              {sz.label} {adjustments[sz.id] > 0 ? `+${adjustments[sz.id]}팩` : `${adjustments[sz.id]}팩`}
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
  const [activeTab,    setActiveTab]    = useState('bandage'); // 'bandage' | 'supply'
  const [stocks,       setStocks]       = useState(defaultStocks);
  const [jeonsan,      setJeonsan]      = useState(defaultJeonsan);
  const [adjustments,  setAdjustments]  = useState(defaultAdj);
  const [ocha,         setOcha]         = useState(defaultOcha);
  const [manualOrders, setManualOrders] = useState(defaultManualOrders);
  const [hydrated,     setHydrated]     = useState(false);
  const [downloading,  setDownloading]  = useState(false);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_KEY);
      if (saved) {
        const p = JSON.parse(saved);
        if (p.stocks)       setStocks(p.stocks);
        if (p.jeonsan)      setJeonsan(p.jeonsan);
        if (p.adjustments)  setAdjustments(p.adjustments);
        if (p.ocha)         setOcha(p.ocha);
        if (p.manualOrders) setManualOrders(p.manualOrders);
        else if (p.stocking) setManualOrders(prev => ({ ...prev, stocking: p.stocking }));
      }
    } catch {}
    setHydrated(true);
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    localStorage.setItem(LS_KEY, JSON.stringify({ stocks, jeonsan, adjustments, ocha, manualOrders }));
  }, [stocks, jeonsan, adjustments, ocha, manualOrders, hydrated]);

  const resetProductAdj = useCallback((productId) => {
    const prod = CALCULATED_PRODUCTS.find(p => p.id === productId);
    setAdjustments(prev => ({
      ...prev,
      [productId]: Object.fromEntries(prod.sizes.map(s => [s.id, 0])),
    }));
  }, []);

  const handleStockChange = useCallback((productId, sizeId, val) => {
    setStocks(prev => ({ ...prev, [productId]: { ...prev[productId], [sizeId]: val } }));
    resetProductAdj(productId);
  }, [resetProductAdj]);

  const handleAdjust = useCallback((productId, sizeId, delta) => {
    setAdjustments(prev => ({
      ...prev,
      [productId]: { ...prev[productId], [sizeId]: (prev[productId][sizeId] || 0) + delta },
    }));
  }, []);

  const handleOchaChange = useCallback((productId, sizeId, val) => {
    setOcha(prev => ({ ...prev, [productId]: { ...prev[productId], [sizeId]: val } }));
  }, []);

  const handleOchaReset = useCallback(() => { setOcha(defaultOcha()); }, []);

  const handleUpload = useCallback((results) => {
    setStocks(prev => {
      const next = { ...prev };
      for (const [productId, jsMap] of Object.entries(results)) {
        next[productId] = { ...next[productId] };
        for (const [sizeId, jsVal] of Object.entries(jsMap)) {
          next[productId][sizeId] = jsVal + (ocha[productId]?.[sizeId] ?? 0);
        }
      }
      return next;
    });
    setJeonsan(prev => {
      const next = { ...prev };
      for (const [productId, jsMap] of Object.entries(results)) {
        next[productId] = { ...next[productId], ...jsMap };
      }
      return next;
    });
    setAdjustments(defaultAdj());
  }, [ocha]);

  const handleManualOrderChange = useCallback((productId, sizeId, val) => {
    setManualOrders(prev => ({ ...prev, [productId]: { ...prev[productId], [sizeId]: val } }));
  }, []);

  const handleDownload = async () => {
    setDownloading(true);
    try {
      const orders = [];
      for (const prod of CALCULATED_PRODUCTS) {
        for (const sz of prod.sizes) {
          const stock = stocks[prod.id][sz.id] ?? sz.default_stock;
          const { packs: base } = calcOrder(sz.avg_monthly, stock);
          const adj = adjustments[prod.id]?.[sz.id] || 0;
          const finalPacks = Math.max(0, base + adj);
          const qty = finalPacks * PACK_SIZE;
          if (qty > 0) orders.push({ product: prod.id, size: sz.id, qty });
        }
      }
      for (const prod of MANUAL_PRODUCTS) {
        for (const sz of prod.sizes) {
          const qty = manualOrders[prod.id]?.[sz.id] || 0;
          if (qty > 0) orders.push({ product: prod.id, size: sz.id, qty });
        }
      }

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
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
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

  const handleReset = () => {
    setStocks(defaultStocks());
    setJeonsan(defaultJeonsan());
    setAdjustments(defaultAdj());
    setOcha(defaultOcha());
    setManualOrders(defaultManualOrders());
  };

  const calcTotalBoxes = (prod) => {
    const packs = prod.sizes.reduce((s, sz) => {
      const stock = stocks[prod.id][sz.id] ?? sz.default_stock;
      const { packs: base } = calcOrder(sz.avg_monthly, stock);
      const adj = adjustments[prod.id]?.[sz.id] || 0;
      return s + Math.max(0, base + adj);
    }, 0);
    return Math.floor(packs / 2);
  };

  const boxesByProduct = Object.fromEntries(CALCULATED_PRODUCTS.map(p => [p.id, calcTotalBoxes(p)]));
  const grandTotal = Object.values(boxesByProduct).reduce((s, b) => s + b, 0);
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
          <h1 className={styles.title}>발주 관리</h1>
          <p className={styles.subtitle}>랩밴드 · 메디켓 · 소모품 발주서 자동 생성</p>
          {activeTab === 'bandage' && (
            <button onClick={handleReset} className={styles.resetBtn}>초기화</button>
          )}
        </header>

        {/* Tab navigation */}
        <div className={styles.tabNav}>
          <button
            className={`${styles.tabBtn} ${activeTab === 'bandage' ? styles.tabBtnActive : ''}`}
            onClick={() => setActiveTab('bandage')}
          >
            💊 랩밴드 / 메디켓 발주
          </button>
          <button
            className={`${styles.tabBtn} ${activeTab === 'supply' ? styles.tabBtnActive : ''}`}
            onClick={() => setActiveTab('supply')}
          >
            📦 소모품 발주
          </button>
        </div>

        <main className={styles.main}>
          {activeTab === 'supply' && <SupplyOrderTab />}
          {activeTab === 'bandage' && <UploadSection onUpload={handleUpload} />}

          {activeTab === 'bandage' && (
            <>
              <OchaPanel ocha={ocha} onChange={handleOchaChange} onReset={handleOchaReset} />

              {CALCULATED_PRODUCTS.map(prod => (
                <ProductTable
                  key={prod.id}
                  product={prod}
                  stocks={stocks[prod.id]}
                  jeonsan={jeonsan[prod.id]}
                  ocha={ocha[prod.id]}
                  adjustments={adjustments[prod.id]}
                  onStockChange={(sizeId, val) => handleStockChange(prod.id, sizeId, val)}
                  onAdjust={(sizeId, delta) => handleAdjust(prod.id, sizeId, delta)}
                  onResetAdj={() => resetProductAdj(prod.id)}
                />
              ))}

              {MANUAL_PRODUCTS.map(prod => (
                <ManualProductSection
                  key={prod.id}
                  product={prod}
                  orders={manualOrders[prod.id] || {}}
                  onChange={handleManualOrderChange}
                />
              ))}

              <div className={`${styles.summary} ${meetsMin ? styles.summaryOk : styles.summaryWarn}`}>
                <div className={styles.summaryRow}>
                  {CALCULATED_PRODUCTS.map((prod, i) => (
                    <Fragment key={prod.id}>
                      {i > 0 && <span className={styles.summaryPlus}>+</span>}
                      <div className={styles.summaryBlock}>
                        <span className={styles.summaryLabel}>{prod.name}</span>
                        <span className={styles.summaryVal}>{boxesByProduct[prod.id]}박스</span>
                      </div>
                    </Fragment>
                  ))}
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
            </>
          )}
        </main>
      </div>
    </>
  );
}
