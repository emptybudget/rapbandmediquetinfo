import Head from 'next/head';
import { useState, useCallback, useEffect, useMemo, Fragment } from 'react';
import {
  PRODUCTS, CALCULATED_PRODUCTS, MANUAL_PRODUCTS, PRODUCT_MAP,
  PACK_SIZE, FREE_SHIP_BOXES,
} from '../lib/products';
import { VENDORS, VENDOR_MAP } from '../lib/vendors';

const SUPPLY_VENDORS = VENDORS.filter(v => v.group === 'supplies');
const SUPPLIES_CARD  = { id: 'supplies', name: '소모품', color: '#3d5a7a', _isGroup: true };
const HOME_VENDORS   = VENDORS.reduce((acc, v) => {
  if (v.group === 'supplies') {
    if (!acc.some(x => x.id === 'supplies')) acc.push(SUPPLIES_CARD);
  } else {
    acc.push(v);
  }
  return acc;
}, []);
import { MEDYSSEY_SEED } from '../lib/medysseyData';
import styles from '../styles/Home.module.css';

// ─── constants ────────────────────────────────────────────────────────────────
const LS_BAND     = 'bandOrderStocks';
const LS_DRAFTS   = 'vendorDrafts';
const LS_REQUESTER = 'lastRequester';
const LS_MED_ADDS = 'medysseyAdds';

// ─── helpers ──────────────────────────────────────────────────────────────────
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

function fmtLogDate(ts) {
  const d = new Date(ts);
  return `${d.getMonth()+1}/${d.getDate()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function fmtLogContent(entry) {
  if (entry.items) {
    const lines = entry.items
      .filter(it => it.qty > 0)
      .map(it => `${it.name}${it.spec ? `(${it.spec})` : ''} ${it.qty}개`)
      .join(', ');
    return entry.hospital ? `[${entry.hospital}] ${lines}` : lines;
  }
  if (entry.tab === 'supply' && entry.items) {
    return entry.items.map(it => `${it.name}${it.spec ? `(${it.spec})` : ''} ${it.qty}개`).join(', ');
  }
  if (entry.orders) {
    const byProduct = {};
    for (const o of entry.orders) {
      (byProduct[o.product] ??= []).push(`${o.size} ${o.qty}개`);
    }
    return Object.entries(byProduct)
      .map(([pid, szs]) => `${PRODUCT_MAP[pid]?.name ?? pid}: ${szs.join(', ')}`)
      .join(' / ');
  }
  return '—';
}

function fmtLogBadge(entry) {
  if (entry.vendor) {
    const v = VENDOR_MAP[entry.vendor];
    return v ? v.name : entry.vendor;
  }
  if (entry.tab === 'bandage') return '💊 랩밴드';
  if (entry.tab === 'supply')  return '📦 소모품';
  return '기타';
}

async function parseExcelStocks(file, product) {
  const XLSX = await import('xlsx');
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const result = {};
  for (const sheetName of wb.SheetNames) {
    const up = sheetName.toUpperCase();
    for (const sz of product.sizes) {
      if (up.includes(`(${sz.id})`)) {
        const cell = wb.Sheets[sheetName]['D3'];
        if (cell != null) result[sz.id] = Number(cell.v);
        break;
      }
    }
  }
  return result;
}

// ─── shared sub-components ────────────────────────────────────────────────────
function SizeChip({ sizeObj, small }) {
  return (
    <span
      className={small ? styles.sizeChipSm : styles.sizeChip}
      style={{ background: sizeObj.chipColor }}
    >{sizeObj.label}</span>
  );
}

// ─── Upload section (repmedicare only) ────────────────────────────────────────
const UPLOAD_PRODUCTS = PRODUCTS.filter(p => p.hasExcelUpload);

function UploadSection({ onUpload }) {
  const [files, setFiles] = useState(() => Object.fromEntries(UPLOAD_PRODUCTS.map(p => [p.id, null])));
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
            <input type="file" accept=".xlsx,.xls" className={styles.fileInput}
              onChange={e => setFiles(prev => ({ ...prev, [prod.id]: e.target.files[0] || null }))} />
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

// ─── Ocha panel ───────────────────────────────────────────────────────────────
function OchaPanel({ ocha, onChange, onReset }) {
  const [open, setOpen] = useState(false);
  return (
    <div className={styles.ochaBox}>
      <button className={styles.ochaToggle} onClick={() => setOpen(v => !v)}>
        <span>⚙️ 오차 설정</span>
        <span className={styles.ochaToggleHint}>{open ? '접기 ▲' : '창고재고 − 전산재고 값 편집 ▼'}</span>
      </button>
      {open && (
        <div className={styles.ochaBody}>
          <p className={styles.ochaDesc}>오차 = 창고 실재고 − 전산 재고량. 엑셀 업로드 시 <strong>창고재고 = 전산 + 오차</strong>로 자동 계산됩니다.</p>
          <div className={styles.ochaGrid}>
            {CALCULATED_PRODUCTS.map(prod => (
              <div key={prod.id} className={styles.ochaProduct}>
                <div className={styles.ochaProductLabel} style={{ borderColor: prod.color }}>{prod.name}</div>
                <div className={styles.ochaRow}>
                  {prod.sizes.map(sz => (
                    <div key={sz.id} className={styles.ochaItem}>
                      <SizeChip sizeObj={sz} small />
                      <input type="number" value={ocha[prod.id]?.[sz.id] ?? 0}
                        onChange={e => onChange(prod.id, sz.id, Number(e.target.value))}
                        className={styles.ochaInput} />
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

// ─── ProductTable ─────────────────────────────────────────────────────────────
function ProductTable({ product, stocks, jeonsan, ocha, adjustments, onStockChange, onAdjust, onResetAdj }) {
  const monthKeys = product.sizes[0]?.monthly ? Object.keys(product.sizes[0].monthly) : [];
  const rows = product.sizes.map(sz => {
    const stock = stocks[sz.id] ?? sz.default_stock;
    const { shortage, packs } = calcOrder(sz.avg_monthly, stock);
    const adjDelta = adjustments[sz.id] || 0;
    const finalPacks = Math.max(0, packs + adjDelta);
    const finalQty = finalPacks * PACK_SIZE;
    const jsVal = jeonsan[sz.id];
    const ochaVal = ocha[sz.id] ?? 0;
    return { ...sz, stock, shortage, basePacks: packs, finalPacks, finalQty, jeonsan: jsVal, ocha: ochaVal };
  });
  const totalPacks = rows.reduce((s, r) => s + r.finalPacks, 0);
  const totalBoxes = Math.floor(totalPacks / 2);
  const isOdd = totalPacks > 0 && totalPacks % 2 !== 0;
  const hasAdj = product.sizes.some(sz => (adjustments[sz.id] || 0) !== 0);
  const orderedRows = rows.filter(r => r.finalPacks > 0);
  const biggestRow = orderedRows.length ? orderedRows.reduce((a, b) => b.finalPacks > a.finalPacks ? b : a) : null;
  const borderlineRow = orderedRows.length ? orderedRows.reduce((a, b) => (a.shortage || 0) <= (b.shortage || 0) ? a : b) : null;

  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: product.color }}>{product.name}</h2>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>사이즈</th><th>창고재고</th><th><span className={styles.thSub}>전산 / 오차</span></th>
              <th>월평균<br />사용량</th><th>적정재고<br />(1배수)</th>
              <th>월별 사용량<br /><span className={styles.small}>{monthKeys.map(k => k.split('-')[1].replace(/^0/,'')+'월').join(' / ')}</span></th>
              <th>부족량</th><th>발주수량<br /><span className={styles.small}>(30개 단위)</span></th>
              <th>팩수</th><th>박스</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(row => {
              const adjDelta = adjustments[row.id] || 0;
              return (
                <tr key={row.id} className={row.finalPacks > 0 ? styles.needOrder : ''}>
                  <td><SizeChip sizeObj={row} /></td>
                  <td><input type="number" min="0" value={row.stock}
                    onChange={e => onStockChange(row.id, Number(e.target.value))} className={styles.stockInput} /></td>
                  <td className={styles.jeonsanCell}>
                    {row.jeonsan != null ? (<>
                      <span className={styles.jeonsanVal}>{row.jeonsan}</span>
                      <span className={row.ocha >= 0 ? styles.deltaPos : styles.deltaNeg}>
                        {row.ocha >= 0 ? `+${row.ocha}` : row.ocha}
                      </span>
                    </>) : <span className={styles.noData}>—</span>}
                  </td>
                  <td className={styles.num}>{row.avg_monthly}</td>
                  <td className={styles.num}>{Math.ceil(row.avg_monthly)}</td>
                  <td className={styles.monthCell}>{monthKeys.map(k => row.monthly[k]).join(' / ')}</td>
                  <td className={`${styles.num} ${row.shortage > 0 ? styles.red : styles.greenTxt}`}>
                    {row.shortage > 0 ? `+${row.shortage}` : '충분'}
                  </td>
                  <td className={`${styles.num} ${row.finalQty > 0 ? styles.bold : ''}`}>
                    {row.finalQty > 0 ? `${row.finalQty}개` : '—'}
                    {adjDelta !== 0 && (
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
                          {row.finalPacks % 2 === 0 ? `${row.finalPacks/2}박스` : `${Math.floor(row.finalPacks/2)}+½`}
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
                <span className={`${styles.boxBadge} ${isOdd ? styles.boxBadgeOdd : ''}`}
                  style={isOdd ? undefined : { background: product.color }}>
                  {isOdd ? `${totalBoxes}+½` : `${totalBoxes}박스`}
                </span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
      {isOdd && (
        <div className={styles.oddPanel}>
          <p className={styles.oddTitle}>⚠️ 총 <strong>{totalPacks}팩</strong> (홀수) — 완성 박스 구성 불가. 조정 방법을 선택하세요:</p>
          <div className={styles.oddBtns}>
            {biggestRow && (
              <button className={styles.oddBtnReduce} onClick={() => onAdjust(biggestRow.id, -1)}>
                <span className={styles.oddBtnLabel}>줄이기</span>
                <SizeChip sizeObj={biggestRow} small />
                <span className={styles.oddBtnDetail}>1팩↓ → {totalPacks-1}팩 = <strong>{(totalPacks-1)/2}박스</strong></span>
              </button>
            )}
            {borderlineRow && (
              <button className={styles.oddBtnAdd} onClick={() => onAdjust(borderlineRow.id, 1)}>
                <span className={styles.oddBtnLabel}>늘리기</span>
                <SizeChip sizeObj={borderlineRow} small />
                <span className={styles.oddBtnDetail}>1팩↑ → {totalPacks+1}팩 = <strong>{(totalPacks+1)/2}박스</strong></span>
              </button>
            )}
          </div>
        </div>
      )}
      {!isOdd && hasAdj && (
        <div className={styles.adjApplied}>
          <span>조정 적용됨:</span>
          {product.sizes.filter(sz => (adjustments[sz.id]||0) !== 0).map(sz => (
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

// ─── ManualProductSection ─────────────────────────────────────────────────────
function ManualProductSection({ product, orders, onChange }) {
  const hasOrder = product.sizes.some(sz => (orders[sz.id] || 0) > 0);
  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: product.color }}>{product.name}</h2>
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
                  <td><input type="number" min="0" value={qty}
                    onChange={e => onChange(product.id, sz.id, Number(e.target.value))}
                    className={styles.stockInput} /></td>
                  <td className={`${styles.num} ${qty > 0 ? styles.bold : ''}`}>{qty > 0 ? `${qty}개` : '—'}</td>
                </tr>
              );
            })}
          </tbody>
          {hasOrder && (
            <tfoot>
              <tr className={styles.totalRow}>
                <td colSpan={2}>총 발주량</td>
                <td className={styles.num}>{product.sizes.reduce((s, sz) => s + (orders[sz.id]||0), 0)}개</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ─── OrderLog ─────────────────────────────────────────────────────────────────
function OrderLog({ log, onClear }) {
  if (!log.length) return null;
  return (
    <div className={styles.section}>
      <h2 className={styles.sectionTitle} style={{ borderLeftColor: '#aaa' }}>
        발주 이력
        <button className={styles.logClearBtn} onClick={onClear}>이력 삭제</button>
      </h2>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th style={{ width: 72, whiteSpace: 'nowrap' }}>일시</th>
              <th style={{ width: 100 }}>구분</th>
              <th>내용</th>
            </tr>
          </thead>
          <tbody>
            {log.map(entry => (
              <tr key={entry.id}>
                <td className={styles.logDate}>{fmtLogDate(entry.ts)}</td>
                <td>
                  <span className={entry.vendor ? styles.logBadgeVendor : (entry.tab === 'bandage' ? styles.logBadgeBandage : styles.logBadgeSupply)}>
                    {fmtLogBadge(entry)}
                  </span>
                </td>
                <td className={styles.logContent}>{fmtLogContent(entry)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Medyssey Tab ─────────────────────────────────────────────────────────────
const MC = {
  paper: '#f4f1e9', ink: '#15302b', sub: '#6a7a74', line: '#e0d9ca',
  card: '#ffffff', accent: '#a8531b', accentSoft: '#f3e6da', good: '#0f766e', userTag: '#3b6ea5',
};
const mSerif = "'Iowan Old Style','Palatino Linotype','Georgia',serif";
const mMono  = "'SFMono-Regular','Menlo','Consolas',monospace";

function mMonthsAgo(last) {
  if (!last) return '';
  const m = Math.floor(((new Date() - new Date(last)) / 86400000) / 30);
  return m <= 0 ? '이번 달 사용' : m + '개월 전 사용';
}

function MTag() {
  return <span style={{ fontSize: 10, fontWeight: 700, color: '#fff', background: MC.userTag, borderRadius: 5, padding: '1px 6px' }}>직접추가</span>;
}

function MAddBox({ value, setValue, onAdd, placeholder, label }) {
  return (
    <div style={{ display: 'flex', gap: 8, marginTop: 14, paddingTop: 14, borderTop: '1px dashed ' + MC.line, maxWidth: 520 }}>
      <div style={{ position: 'relative', flex: 1 }}>
        <span style={{ position: 'absolute', left: 11, top: 13, color: MC.sub, fontSize: 14 }}>✏️</span>
        <input value={value} onChange={e => setValue(e.target.value)} onKeyDown={e => e.key === 'Enter' && onAdd()} placeholder={placeholder}
          style={{ width: '100%', padding: '11px 12px 11px 34px', borderRadius: 11, border: '1px solid ' + MC.line, background: MC.card, fontSize: 14, boxSizing: 'border-box' }} />
      </div>
      <button onClick={onAdd} style={{ display: 'flex', alignItems: 'center', gap: 5, padding: '0 16px', borderRadius: 11, border: 'none', background: MC.userTag, color: '#fff', fontWeight: 700, cursor: 'pointer', fontSize: 14, whiteSpace: 'nowrap' }}>
        + {label}
      </button>
    </div>
  );
}

function MCrumb({ n, label, active, done, dim, onClick }) {
  return (
    <button onClick={onClick} disabled={dim} style={{ display: 'flex', alignItems: 'center', gap: 6, border: 'none', background: 'none', cursor: dim ? 'default' : 'pointer', color: active ? MC.ink : (dim ? '#bcc4c0' : MC.sub), fontSize: 13, fontWeight: active ? 700 : 500, padding: 0, maxWidth: 160, overflow: 'hidden' }}>
      <span style={{ width: 20, height: 20, borderRadius: 20, background: active ? MC.accent : (done ? MC.good : 'transparent'), border: (active || done) ? 'none' : '1px solid ' + MC.line, color: (active || done) ? '#fff' : MC.sub, fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, fontFamily: mMono }}>{n}</span>
      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</span>
    </button>
  );
}

function MSection({ guide, children }) {
  return (
    <div style={{ paddingBottom: 20 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '4px 0 14px', fontSize: 15, fontWeight: 700 }}>
        <span style={{ color: MC.accent }}>›</span>{guide}
      </div>
      {children}
    </div>
  );
}

function MEmpty({ children }) {
  return <div style={{ color: '#9aa39f', fontSize: 14, textAlign: 'center', padding: '24px 0' }}>{children}</div>;
}

function MBigStepper({ qty, onSet }) {
  const btn = (on) => ({
    width: 38, height: 38, borderRadius: 10, border: '1px solid ' + MC.line,
    background: on ? MC.ink : MC.card, color: on ? MC.paper : '#c4ccc8',
    cursor: on ? 'pointer' : 'default', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
  });
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
      <button style={btn(qty > 0)} onClick={() => onSet(Math.max(0, qty - 1))}>−</button>
      <span style={{ fontFamily: mMono, fontWeight: 700, fontSize: 18, minWidth: 26, textAlign: 'center', color: qty > 0 ? MC.ink : '#c4ccc8' }}>{qty}</span>
      <button style={btn(true)} onClick={() => onSet(qty + 1)}>+</button>
    </div>
  );
}

// 볼트·로드·커넥터류는 기구 목록에서 제외하고 스크루 화면 내에 표시
function isAccessoryInstrument(name) {
  return /bolt|rod|connector/i.test(name) && !/screw|poly|mono|iliac/i.test(name);
}

// 선택된 스크루 계통에 맞는 부속품 반환; 커넥터 종류는 하나로 병합, MIS 로드는 types 구조로 반환
function getRelevantAccessories(instName, hospitalData, userSz, hosp) {
  const isMIS    = /mis/i.test(instName);
  const isZenius = /zenius/i.test(instName);
  const isILIAD  = /iliad/i.test(instName);

  // merged: displayName → { sizes, types, w }
  // types 있으면 [{label, key, sizes}] — rod 다중 타입용
  const merged = new Map();

  for (const [name, v] of Object.entries(hospitalData)) {
    if (!isAccessoryInstrument(name)) continue;

    const isConn = /connector|transverse/i.test(name);
    const isRod  = /rod/i.test(name) && !isConn;

    if (isILIAD) {
      if (!isConn && !/iliad/i.test(name)) continue;
    } else if (isZenius) {
      if (/iliad/i.test(name)) continue;
    }
    if (isMIS && isRod && !/mis/i.test(name)) continue;
    // 볼트는 "Bolt" / "ILIAD Bolt" 하나만 사용, 나머지 변형 제외
    if (/bolt/i.test(name) && !/^(iliad\s+)?bolt$/i.test(name)) continue;

    let displayName, typeLabel = null;
    if (isConn) {
      displayName = 'Rod Connector';
    } else if (isMIS && isRod) {
      displayName = 'Rod';
      typeLabel   = /cov/i.test(name) ? 'Curved' : 'Straight';
    } else {
      displayName = name;
    }

    if (!merged.has(displayName)) merged.set(displayName, { sizes: [], types: null, w: 0 });
    const entry = merged.get(displayName);
    entry.w = Math.max(entry.w, v.w);

    if (typeLabel) {
      if (!entry.types) entry.types = [];
      let t = entry.types.find(x => x.label === typeLabel);
      if (!t) { t = { label: typeLabel, key: name, sizes: [] }; entry.types.push(t); }
      for (const s of v.sizes) if (!t.sizes.some(e => e.size === s.size)) t.sizes.push(s);
    } else {
      for (const s of v.sizes) if (!entry.sizes.some(e => e.size === s.size)) entry.sizes.push(s);
    }
  }

  return [...merged.entries()].map(([name, entry]) => {
    if (entry.types) {
      const types = entry.types
        .sort((a, b) => a.label === 'Curved' ? -1 : b.label === 'Curved' ? 1 : 0)
        .map(t => {
          const us = userSz[hosp + '||' + t.key] || [];
          const usSet = new Set(us);
          const extra = us.map(s => ({ size: s, last: null, _user: true }));
          const sizes = [...extra, ...t.sizes.filter(s => !usSet.has(s.size))]
            .sort((a, b) => (a.size||'').localeCompare(b.size||'', undefined, { numeric: true, sensitivity: 'base' }));
          return { ...t, sizes };
        });
      return { name, types, w: entry.w };
    }
    const us = userSz[hosp + '||' + name] || [];
    const usSet = new Set(us);
    const extra = us.map(s => ({ size: s, last: null, _user: true }));
    const sizes = [...extra, ...entry.sizes.filter(s => !usSet.has(s.size))]
      .sort((a, b) => (a.size||'').localeCompare(b.size||'', undefined, { numeric: true, sensitivity: 'base' }));
    return { name, sizes, w: entry.w };
  }).sort((a, b) => b.w - a.w);
}

function MedysseyTab({ adds, onAddInst, onAddSize, onRemoveSize, cart, onCartChange, onDownload, downloading, requester }) {
  const today = new Date().toISOString().slice(0, 10);

  const hospitals = useMemo(() => {
    const hospMap = {};
    for (const [h, insts] of Object.entries(MEDYSSEY_SEED.hospitals)) {
      hospMap[h] = { name: h, instCount: Object.keys(insts).length, w: Object.values(insts).reduce((a, i) => a + i.w, 0) };
    }
    // Add user-added hospitals
    for (const key of Object.keys(adds)) {
      const [h] = key.split('|');
      if (h && !hospMap[h]) hospMap[h] = { name: h, instCount: 0, w: 0, _user: true };
    }
    return Object.values(hospMap).sort((a, b) => b.w - a.w);
  }, [adds]);

  const [step, setStep] = useState(0);
  const [hosp, setHosp] = useState(null);
  const [inst, setInst] = useState(null);
  const [q, setQ] = useState('');
  const [sheet, setSheet] = useState(false);
  const [copied, setCopied] = useState(false);
  const [newInst, setNewInst] = useState('');
  const [newSizes, setNewSizes] = useState({});
  const [remarks, setRemarks] = useState('');
  const [accTypeSelect, setAccTypeSelect] = useState({});

  // Reconstruct userInst and userSize from flat adds object
  const { userInst, userSize } = useMemo(() => {
    const ui = {};
    const us = {};
    for (const key of Object.keys(adds)) {
      const parts = key.split('|');
      const h = parts[0], i = parts[1], s = parts[2];
      if (!h || !i) continue;
      if (!ui[h]) ui[h] = [];
      if (!ui[h].includes(i)) ui[h].push(i);
      if (s) {
        const k = h + '||' + i;
        if (!us[k]) us[k] = [];
        if (!us[k].includes(s)) us[k].push(s);
      }
    }
    return { userInst: ui, userSize: us };
  }, [adds]);

  const instList = useMemo(() => {
    if (!hosp) return [];
    const baseData = MEDYSSEY_SEED.hospitals[hosp] || {};
    const base = Object.entries(baseData).filter(([name]) => !isAccessoryInstrument(name)).map(([name, v]) => ({ name, sizes: v.sizes, w: v.w }));
    const baseNames = new Set(base.map(b => b.name));
    const extra = (userInst[hosp] || []).filter(n => !baseNames.has(n) && !isAccessoryInstrument(n)).map(name => ({ name, sizes: [], w: Infinity, _user: true }));
    return [...extra, ...base].filter(i => i.name.toLowerCase().includes(q.toLowerCase())).sort((a, b) => b.w - a.w);
  }, [hosp, q, userInst]);

  const sizes = useMemo(() => {
    if (!inst || !hosp) return [];
    const baseData = MEDYSSEY_SEED.hospitals[hosp]?.[inst];
    const base = baseData ? baseData.sizes : [];
    const us = userSize[hosp + '||' + inst] || [];
    const usSet = new Set(us);
    const extra = us.map(size => ({ size, last: null, _user: true }));
    const all = [...extra, ...base.filter(s => !usSet.has(s.size))];
    return all.sort((a, b) => (a.size || '').localeCompare(b.size || '', undefined, { numeric: true, sensitivity: 'base' }));
  }, [hosp, inst, userSize]);

  const accessories = useMemo(() => {
    if (!inst || !hosp) return [];
    return getRelevantAccessories(inst, MEDYSSEY_SEED.hospitals[hosp] || {}, userSize, hosp);
  }, [inst, hosp, userSize]);

  const keyOf = (i, s) => i + '||' + s;
  const getQty  = (i, s) => cart[keyOf(i, s)]?.qty || 0;
  const getNote = (i, s) => cart[keyOf(i, s)]?.note || '';
  const setQty  = (i, s, n) => onCartChange(prev => {
    const k = keyOf(i, s);
    if (n <= 0) { const x = { ...prev }; delete x[k]; return x; }
    return { ...prev, [k]: { ...(prev[k] || { inst: i, size: s }), qty: n } };
  });
  const setNote = (i, s, note) => onCartChange(prev => {
    const k = keyOf(i, s);
    if (!prev[k]) return prev;
    return { ...prev, [k]: { ...prev[k], note } };
  });

  const cartArr = Object.values(cart);
  const totalQty = cartArr.reduce((a, i) => a + i.qty, 0);

  const copyTSV = () => {
    const txt = 'NO\t품목\t규격\t수량\n' + cartArr.map((it, x) => `${x+1}\t${it.inst}\t${it.size||''}\t${it.qty}`).join('\n');
    try { navigator.clipboard.writeText(txt); } catch {}
    setCopied(true); setTimeout(() => setCopied(false), 1500);
  };

  const goHosp = h => { setHosp(h); setInst(null); setQ(''); setStep(1); };
  const goInst = i => { setInst(i); setNewSize(''); setStep(2); };

  const addInst = () => {
    const n = newInst.trim();
    if (!n) return;
    onAddInst(hosp, n);
    setNewInst('');
    goInst(n);
  };

  const addSizeFor = (instKey) => {
    const s = (newSizes[instKey] || '').trim();
    if (!s) return;
    onAddSize(hosp, instKey, s);
    setNewSizes(prev => ({ ...prev, [instKey]: '' }));
    setQty(instKey, s, 1);
  };

  return (
    <div style={{ background: MC.paper, color: MC.ink, fontFamily: 'system-ui,-apple-system,sans-serif', paddingBottom: 90, minHeight: '60vh' }}>
      {/* Breadcrumb */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '12px 20px', fontSize: 13, color: MC.sub, flexWrap: 'wrap', borderBottom: '1px solid ' + MC.line }}>
        <MCrumb n="1" label={hosp || '병원'} active={step === 0} done={step > 0} onClick={() => setStep(0)} />
        <span style={{ opacity: .4 }}>›</span>
        <MCrumb n="2" label={inst || '기구'} active={step === 1} done={step > 1} dim={!hosp} onClick={() => hosp && setStep(1)} />
        <span style={{ opacity: .4 }}>›</span>
        <MCrumb n="3" label="사양 / 수량" active={step === 2} dim={!inst} onClick={() => inst && setStep(2)} />
      </div>

      <div style={{ padding: '0 20px' }}>
        {/* Step 0: Hospital */}
        {step === 0 && (
          <MSection guide="① 병원을 고르세요">
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(150px,1fr))', gap: 10 }}>
              {hospitals.map(h => (
                <button key={h.name} onClick={() => goHosp(h.name)} style={{ textAlign: 'left', padding: 16, borderRadius: 13, border: '1px solid ' + MC.line, background: MC.card, cursor: 'pointer' }}>
                  <div style={{ fontFamily: mSerif, fontSize: 19 }}>{h.name}</div>
                  <div style={{ fontSize: 12, color: MC.sub, marginTop: 4 }}>
                    {h._user ? '직접추가' : `기구 ${h.instCount}종`}
                  </div>
                </button>
              ))}
            </div>
          </MSection>
        )}

        {/* Step 1: Instrument */}
        {step === 1 && hosp && (
          <MSection guide="② 기구를 고르세요">
            <div style={{ position: 'relative', marginBottom: 12, maxWidth: 420 }}>
              <span style={{ position: 'absolute', left: 11, top: 13, color: MC.sub }}>🔍</span>
              <input autoFocus value={q} onChange={e => setQ(e.target.value)} placeholder="기구 이름 검색"
                style={{ width: '100%', padding: '11px 12px 11px 34px', borderRadius: 11, border: '1px solid ' + MC.line, background: MC.card, fontSize: 14, boxSizing: 'border-box' }} />
            </div>
            <div style={{ display: 'grid', gap: 8 }}>
              {instList.map(i => {
                const inOrder = cartArr.filter(c => c.inst === i.name).reduce((a, c) => a + c.qty, 0);
                return (
                  <button key={i.name} onClick={() => goInst(i.name)} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '15px 16px', borderRadius: 12, border: '1px solid ' + MC.line, background: MC.card, cursor: 'pointer' }}>
                    <span style={{ fontWeight: 600, fontSize: 15, display: 'flex', alignItems: 'center', gap: 8 }}>
                      {i.name}{i._user && <MTag />}
                    </span>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 10, color: MC.sub, fontSize: 12 }}>
                      {inOrder > 0 && <span style={{ background: MC.accent, color: '#fff', borderRadius: 20, padding: '2px 9px', fontWeight: 700, fontFamily: mMono }}>{inOrder}</span>}
                      <span>{i.sizes.length}개 사양</span>
                      <span>›</span>
                    </span>
                  </button>
                );
              })}
              {instList.length === 0 && <MEmpty>목록에 없어요 — 아래에서 직접 추가하세요</MEmpty>}
            </div>
            <MAddBox value={newInst} setValue={setNewInst} onAdd={addInst} placeholder="목록에 없는 기구 직접 입력" label="기구 추가" />
          </MSection>
        )}

        {/* Step 2: Size/Qty */}
        {step === 2 && inst && (
          <MSection guide="③ 사양 옆 −／＋ 로 수량을 정하세요">
            <div style={{ display: 'grid', gap: 9 }}>
              {sizes.map(s => {
                const qty  = getQty(inst, s.size);
                const note = getNote(inst, s.size);
                return (
                  <div key={s.size || '_'} style={{ borderRadius: 12, border: '1px solid ' + (qty > 0 ? MC.accent : MC.line), background: qty > 0 ? MC.accentSoft : MC.card, overflow: 'hidden' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, padding: '13px 14px' }}>
                      <div style={{ minWidth: 0, flex: 1 }}>
                        <div style={{ fontFamily: mMono, fontWeight: 700, fontSize: 16, display: 'flex', alignItems: 'center', gap: 7 }}>
                          <span>{s.size || '단일 (규격 없음)'}</span>
                          {s._user && <MTag />}
                          {s._user && (
                            <button onClick={() => onRemoveSize(hosp, inst, s.size)}
                              style={{ padding: '1px 6px', borderRadius: 5, border: '1px solid #e0362088', background: 'transparent', color: '#c0392b', fontSize: 12, cursor: 'pointer', lineHeight: 1.4 }}>×</button>
                          )}
                        </div>
                        {mMonthsAgo(s.last) && <div style={{ fontSize: 11, color: MC.sub, marginTop: 2 }}>{mMonthsAgo(s.last)}</div>}
                      </div>
                      <MBigStepper qty={qty} onSet={n => setQty(inst, s.size, n)} />
                    </div>
                    {qty > 0 && (
                      <div style={{ padding: '0 14px 10px' }}>
                        <input value={note} onChange={e => setNote(inst, s.size, e.target.value)}
                          placeholder="비고 (오픈·반품·교환 등)"
                          style={{ width: '100%', padding: '5px 8px', borderRadius: 6, border: '1px solid ' + MC.line, fontSize: 12, background: 'transparent', boxSizing: 'border-box', fontFamily: 'inherit', color: MC.ink }} />
                      </div>
                    )}
                  </div>
                );
              })}
              {sizes.length === 0 && <MEmpty>사양이 없어요 — 아래에서 직접 추가하세요</MEmpty>}
            </div>
            <MAddBox value={newSizes[inst] || ''} setValue={v => setNewSizes(p => ({ ...p, [inst]: v }))} onAdd={() => addSizeFor(inst)} placeholder="목록에 없는 사양 직접 입력 (예: 6.5*55)" label="사양 추가" />
            {accessories.length > 0 && (
              <div style={{ marginTop: 20, borderTop: '2px dashed ' + MC.line, paddingTop: 16 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: MC.sub, marginBottom: 12 }}>🔩 함께 발주</div>
                {accessories.map(acc => {
                  const renderSizeRow = (instKey, s) => {
                    const qty  = getQty(instKey, s.size);
                    const note = getNote(instKey, s.size);
                    return (
                      <div key={s.size || '_'} style={{ borderRadius: 10, border: '1px solid ' + (qty > 0 ? MC.accent : MC.line), background: qty > 0 ? MC.accentSoft : MC.card, overflow: 'hidden' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, padding: '10px 12px' }}>
                          <div>
                            <div style={{ fontFamily: mMono, fontWeight: 700, fontSize: 14 }}>{s.size || '단일'}</div>
                            {mMonthsAgo(s.last) && <div style={{ fontSize: 11, color: MC.sub }}>{mMonthsAgo(s.last)}</div>}
                          </div>
                          <MBigStepper qty={qty} onSet={n => setQty(instKey, s.size, n)} />
                        </div>
                        {qty > 0 && (
                          <div style={{ padding: '0 12px 8px' }}>
                            <input value={note} onChange={e => setNote(instKey, s.size, e.target.value)}
                              placeholder="비고 (오픈·반품·교환 등)"
                              style={{ width: '100%', padding: '4px 8px', borderRadius: 6, border: '1px solid ' + MC.line, fontSize: 11, background: 'transparent', boxSizing: 'border-box', fontFamily: 'inherit', color: MC.ink }} />
                          </div>
                        )}
                      </div>
                    );
                  };

                  if (acc.types) {
                    const selLabel = accTypeSelect[acc.name] ?? acc.types[0].label;
                    const selType  = acc.types.find(t => t.label === selLabel) ?? acc.types[0];
                    return (
                      <div key={acc.name} style={{ marginBottom: 14 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: MC.ink, marginBottom: 8, fontFamily: mMono }}>{acc.name}</div>
                        <div style={{ display: 'flex', gap: 6, marginBottom: 10 }}>
                          {acc.types.map(t => (
                            <button key={t.label}
                              onClick={() => setAccTypeSelect(s => ({ ...s, [acc.name]: t.label }))}
                              style={{ padding: '6px 16px', borderRadius: 8, border: '1.5px solid ' + (selLabel === t.label ? MC.accent : MC.line), background: selLabel === t.label ? MC.accentSoft : MC.card, color: selLabel === t.label ? MC.accent : MC.ink, fontWeight: selLabel === t.label ? 700 : 400, fontSize: 13, cursor: 'pointer' }}>
                              {t.label}
                            </button>
                          ))}
                        </div>
                        <div style={{ display: 'grid', gap: 7 }}>
                          {selType.sizes.map(s => renderSizeRow(selType.key, s))}
                        </div>
                        <MAddBox value={newSizes[selType.key] || ''} setValue={v => setNewSizes(p => ({ ...p, [selType.key]: v }))} onAdd={() => addSizeFor(selType.key)} placeholder="사양 직접 입력" label="사양 추가" />
                      </div>
                    );
                  }

                  const isBoltAcc = /^(iliad\s+)?bolt$/i.test(acc.name);
                  return (
                    <div key={acc.name} style={{ marginBottom: 14 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: MC.ink, marginBottom: 6, fontFamily: mMono }}>{acc.name}</div>
                      <div style={{ display: 'grid', gap: 7 }}>
                        {acc.sizes.map(s => renderSizeRow(acc.name, s))}
                      </div>
                      {!isBoltAcc && (
                        <MAddBox value={newSizes[acc.name] || ''} setValue={v => setNewSizes(p => ({ ...p, [acc.name]: v }))} onAdd={() => addSizeFor(acc.name)} placeholder="사양 직접 입력" label="사양 추가" />
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </MSection>
        )}
      </div>

      {/* Sticky cart bar */}
      {totalQty > 0 && (
        <div style={{ position: 'fixed', left: 0, right: 0, bottom: 0, background: MC.ink, color: MC.paper, padding: '13px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', boxShadow: '0 -6px 20px rgba(0,0,0,.12)', zIndex: 5 }}>
          <div style={{ fontSize: 14 }}><b style={{ fontFamily: mMono }}>{cartArr.length}</b> 품목 · 총 <b style={{ fontFamily: mMono }}>{totalQty}</b> 개</div>
          <button onClick={() => setSheet(true)} style={{ display: 'flex', alignItems: 'center', gap: 7, background: MC.accent, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 16px', fontWeight: 700, cursor: 'pointer', fontSize: 14 }}>
            📋 발주서 보기
          </button>
        </div>
      )}

      {/* Modal */}
      {sheet && (
        <div onClick={() => setSheet(false)} style={{ position: 'fixed', inset: 0, background: 'rgba(20,30,28,.45)', display: 'flex', alignItems: 'flex-end', justifyContent: 'center', zIndex: 10 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: MC.paper, width: '100%', maxWidth: 640, borderRadius: '18px 18px 0 0', maxHeight: '82vh', display: 'flex', flexDirection: 'column' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid ' + MC.line }}>
              <div style={{ fontFamily: mSerif, fontSize: 20 }}>
                발주서 <span style={{ fontSize: 13, color: MC.sub, fontFamily: mMono }}>· {cartArr.length}품목 / {totalQty}개</span>
              </div>
              <button onClick={() => setSheet(false)} style={{ border: 'none', background: 'none', cursor: 'pointer', color: MC.sub, fontSize: 20 }}>✕</button>
            </div>
            <div style={{ overflowY: 'auto', padding: '6px 12px' }}>
              {cartArr.map((it, idx) => (
                <div key={it.inst + it.size} style={{ display: 'grid', gridTemplateColumns: '24px 1fr auto', alignItems: 'center', gap: 10, padding: '10px 8px', borderBottom: '1px solid ' + MC.line }}>
                  <div style={{ fontFamily: mMono, color: MC.sub, fontSize: 13 }}>{idx + 1}</div>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{it.inst}</div>
                    <div style={{ fontFamily: mMono, fontSize: 12, color: MC.sub }}>{it.size || '단일'}</div>
                    {it.note && <div style={{ fontSize: 11, color: MC.sub, marginTop: 2, fontStyle: 'italic' }}>{it.note}</div>}
                  </div>
                  <MBigStepper qty={it.qty} onSet={n => setQty(it.inst, it.size, n)} />
                </div>
              ))}
            </div>
            <div style={{ padding: '12px 20px', borderTop: '1px solid ' + MC.line }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: MC.sub, marginBottom: 7 }}>비 고</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 8 }}>
                {['방문수령합니다', '택배수령합니다'].map(t => (
                  <button key={t} onClick={() => setRemarks(r => r === t ? '' : t)}
                    style={{ padding: '4px 11px', borderRadius: 20, border: '1px solid ' + (remarks === t ? MC.accent : MC.line), background: remarks === t ? MC.accentSoft : MC.card, color: remarks === t ? MC.accent : MC.ink, fontSize: 12, cursor: 'pointer', fontWeight: remarks === t ? 700 : 400 }}>
                    {t}
                  </button>
                ))}
              </div>
              <textarea value={remarks} onChange={e => setRemarks(e.target.value)}
                placeholder="비고 내용 직접 입력 (예: 단기가납 SET / 홍길동 교수 / 수술 7/10)"
                rows={2} style={{ width: '100%', padding: '8px 10px', borderRadius: 8, border: '1px solid ' + MC.line, fontSize: 13, resize: 'vertical', boxSizing: 'border-box', fontFamily: 'inherit', background: MC.card }} />
            </div>
            <div style={{ display: 'flex', gap: 10, padding: '14px 20px', borderTop: '1px solid ' + MC.line }}>
              <button onClick={() => onCartChange({})} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '11px 14px', borderRadius: 10, border: '1px solid ' + MC.line, background: MC.card, color: MC.sub, cursor: 'pointer', fontSize: 13 }}>
                🗑 비우기
              </button>
              <button onClick={copyTSV} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 7, padding: '11px 14px', borderRadius: 10, border: '1px solid ' + MC.line, background: MC.card, fontWeight: 600, cursor: 'pointer', fontSize: 14 }}>
                {copied ? '✅ 복사됐어요' : '📋 표 복사'}
              </button>
              <button onClick={() => { setSheet(false); onDownload(hosp, cartArr, remarks); }} disabled={downloading}
                style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 7, padding: '11px 14px', borderRadius: 10, border: 'none', background: MC.accent, color: '#fff', fontWeight: 700, cursor: downloading ? 'default' : 'pointer', fontSize: 14, opacity: downloading ? 0.6 : 1 }}>
                {downloading ? '생성 중...' : '📄 엑셀 다운로드'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Vendor home card ─────────────────────────────────────────────────────────
function VendorCard({ vendor, onClick }) {
  if (vendor._isGroup) {
    return (
      <button className={styles.vendorCard} onClick={onClick}>
        <div className={styles.vendorCardColor} style={{ background: vendor.color }}>📦</div>
        <p className={styles.vendorCardName}>{vendor.name}</p>
        <p className={styles.vendorCardMeta}>{SUPPLY_VENDORS.map(v => v.name).join(' · ')}</p>
      </button>
    );
  }
  const icon = vendor.type === 'history' ? '🔬'
    : vendor.products.some(p => p.type === 'calculated') ? '💊' : '📦';
  const meta = vendor.type === 'history'
    ? '병원별 기구 발주'
    : vendor.products.map(p => p.name).join(' · ');

  return (
    <button className={styles.vendorCard} onClick={onClick}>
      <div className={styles.vendorCardColor} style={{ background: vendor.color }}>{icon}</div>
      <p className={styles.vendorCardName}>{vendor.name}</p>
      <p className={styles.vendorCardMeta}>{meta}</p>
    </button>
  );
}

// ─── Main page ─────────────────────────────────────────────────────────────────
export default function Home() {
  const [activeVendorId, setActiveVendorId] = useState(null);
  const [activeSupplyVendorId, setActiveSupplyVendorId] = useState(null);
  const [hydrated, setHydrated] = useState(false);
  const [orderLog, setOrderLog] = useState([]);
  const [requester, setRequester] = useState('');

  // Repmedicare state
  const [stocks, setStocks]           = useState(defaultStocks);
  const [jeonsan, setJeonsan]         = useState(defaultJeonsan);
  const [adjustments, setAdjustments] = useState(defaultAdj);
  const [ocha, setOcha]               = useState(defaultOcha);
  const [manualOrders, setManualOrders] = useState(defaultManualOrders);

  // Other vendor manual orders: { [vendorId]: { [productId]: { [sizeId]: qty } } }
  const [vendorDrafts, setVendorDrafts] = useState({});

  // Medyssey
  const [medysseyAdds, setMedysseyAdds] = useState({});
  const [medysseyCart, setMedysseyCart] = useState({});
  const [medDownloading, setMedDownloading] = useState(false);

  // Repmedicare download state
  const [downloading, setDownloading] = useState(false);

  // ── Hydration (localStorage + server fetch) ───────────────────────────────
  useEffect(() => {
    try {
      const saved = localStorage.getItem(LS_BAND);
      if (saved) {
        const p = JSON.parse(saved);
        if (p.stocks)       setStocks(p.stocks);
        if (p.jeonsan)      setJeonsan(p.jeonsan);
        if (p.adjustments)  setAdjustments(p.adjustments);
        if (p.ocha)         setOcha(p.ocha);
        if (p.manualOrders) setManualOrders(p.manualOrders);
      }
    } catch {}
    try {
      const dr = localStorage.getItem(LS_DRAFTS);
      if (dr) setVendorDrafts(JSON.parse(dr));
    } catch {}
    try {
      const req = localStorage.getItem(LS_REQUESTER);
      if (req) setRequester(req);
    } catch {}

    // Medyssey adds: load from localStorage first, then merge with server
    const localAdds = (() => {
      try { return JSON.parse(localStorage.getItem(LS_MED_ADDS) || '{}'); } catch { return {}; }
    })();
    fetch('/api/medyssey-adds')
      .then(r => r.json())
      .then(serverAdds => {
        const merged = { ...localAdds, ...serverAdds };
        setMedysseyAdds(merged);
        try { localStorage.setItem(LS_MED_ADDS, JSON.stringify(merged)); } catch {}
      })
      .catch(() => setMedysseyAdds(localAdds));

    setHydrated(true);

    // Shared log
    fetch('/api/log')
      .then(r => r.json())
      .then(data => { if (Array.isArray(data)) setOrderLog(data); })
      .catch(() => {});
  }, []);

  // ── Persist repmedicare state ─────────────────────────────────────────────
  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(LS_BAND, JSON.stringify({ stocks, jeonsan, adjustments, ocha, manualOrders })); } catch {}
  }, [stocks, jeonsan, adjustments, ocha, manualOrders, hydrated]);

  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(LS_DRAFTS, JSON.stringify(vendorDrafts)); } catch {}
  }, [vendorDrafts, hydrated]);

  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(LS_REQUESTER, requester); } catch {}
  }, [requester, hydrated]);

  // ── Handlers ──────────────────────────────────────────────────────────────
  const resetProductAdj = useCallback((productId) => {
    const prod = CALCULATED_PRODUCTS.find(p => p.id === productId);
    setAdjustments(prev => ({ ...prev, [productId]: Object.fromEntries(prod.sizes.map(s => [s.id, 0])) }));
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

  const handleOchaReset = useCallback(() => setOcha(defaultOcha()), []);

  const handleUpload = useCallback((results) => {
    setStocks(prev => {
      const next = { ...prev };
      for (const [pid, jsMap] of Object.entries(results)) {
        next[pid] = { ...next[pid] };
        for (const [sid, jsVal] of Object.entries(jsMap)) {
          next[pid][sid] = jsVal + (ocha[pid]?.[sid] ?? 0);
        }
      }
      return next;
    });
    setJeonsan(prev => {
      const next = { ...prev };
      for (const [pid, jsMap] of Object.entries(results)) next[pid] = { ...next[pid], ...jsMap };
      return next;
    });
    setAdjustments(defaultAdj());
  }, [ocha]);

  const handleManualOrderChange = useCallback((productId, sizeId, val) => {
    setManualOrders(prev => ({ ...prev, [productId]: { ...prev[productId], [sizeId]: val } }));
  }, []);

  const handleVendorDraftChange = useCallback((vendorId, productId, sizeId, val) => {
    setVendorDrafts(prev => ({
      ...prev,
      [vendorId]: { ...(prev[vendorId] || {}), [productId]: { ...(prev[vendorId]?.[productId] || {}), [sizeId]: val } },
    }));
  }, []);

  const addLogEntry = useCallback((entry) => {
    const newEntry = { ...entry, id: Date.now(), ts: new Date().toISOString() };
    setOrderLog(prev => [newEntry, ...prev].slice(0, 200));
    fetch('/api/log', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newEntry),
    }).catch(() => {});
  }, []);

  // ── Medyssey adds ─────────────────────────────────────────────────────────
  const handleAddInst = useCallback((hospital, instrument) => {
    const key = `${hospital}|${instrument}|`;
    setMedysseyAdds(prev => {
      const next = { ...prev, [key]: '1' };
      try { localStorage.setItem(LS_MED_ADDS, JSON.stringify(next)); } catch {}
      return next;
    });
    fetch('/api/medyssey-adds', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ hospital, instrument, size: '' }),
    }).catch(() => {});
  }, []);

  const handleAddSize = useCallback((hospital, instrument, size) => {
    const key = `${hospital}|${instrument}|${size}`;
    setMedysseyAdds(prev => {
      const next = { ...prev, [key]: '1' };
      try { localStorage.setItem(LS_MED_ADDS, JSON.stringify(next)); } catch {}
      return next;
    });
    fetch('/api/medyssey-adds', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ hospital, instrument, size }),
    }).catch(() => {});
  }, []);

  const handleRemoveSize = useCallback((hospital, instrument, size) => {
    const key = `${hospital}|${instrument}|${size}`;
    setMedysseyAdds(prev => {
      const next = { ...prev };
      delete next[key];
      try { localStorage.setItem(LS_MED_ADDS, JSON.stringify(next)); } catch {}
      return next;
    });
    fetch('/api/medyssey-adds', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ hospital, instrument, size }),
    }).catch(() => {});
  }, []);

  // ── Medyssey download ─────────────────────────────────────────────────────
  const handleMedysseyDownload = useCallback(async (hospital, cartArr, note) => {
    setMedDownloading(true);
    try {
      const items = cartArr.map(c => ({ name: c.inst, spec: c.size, qty: c.qty, note: c.note || undefined }));
      const res = await fetch('/api/generate-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ vendorId: 'medyssey', requester, items, hospital, note }),
      });
      if (!res.ok) throw new Error('서버 오류 ' + res.status);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const today = new Date();
      const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
      a.href = url;
      a.download = `${hospital ? hospital + '_' : ''}메디쎄이_발주서_${ds}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      addLogEntry({ vendor: 'medyssey', tab: 'medyssey', hospital, items });
    } catch (e) {
      alert('다운로드 실패: ' + e.message);
    }
    setMedDownloading(false);
  }, [requester, addLogEntry]);

  // ── Repmedicare download ──────────────────────────────────────────────────
  const calcTotalBoxes = (prod) => {
    const packs = prod.sizes.reduce((s, sz) => {
      const stock = stocks[prod.id]?.[sz.id] ?? sz.default_stock;
      const { packs: base } = calcOrder(sz.avg_monthly, stock);
      const adj = adjustments[prod.id]?.[sz.id] || 0;
      return s + Math.max(0, base + adj);
    }, 0);
    return Math.floor(packs / 2);
  };

  const handleVendorDownload = async (vendor) => {
    setDownloading(true);
    try {
      const items = [];
      for (const prod of vendor.products) {
        if (prod.type === 'calculated') {
          for (const sz of prod.sizes) {
            const stock = stocks[prod.id]?.[sz.id] ?? sz.default_stock;
            const { packs: base } = calcOrder(sz.avg_monthly, stock);
            const adj = adjustments[prod.id]?.[sz.id] || 0;
            const finalPacks = Math.max(0, base + adj);
            if (finalPacks > 0) items.push({ name: prod.excelName, spec: sz.excelSize, qty: finalPacks * PACK_SIZE });
          }
        } else {
          const orders = vendor.id === 'repmedicare' ? manualOrders : vendorDrafts[vendor.id] || {};
          for (const sz of prod.sizes) {
            const qty = orders[prod.id]?.[sz.id] || 0;
            if (qty > 0) items.push({ name: prod.excelName, spec: sz.excelSize, qty });
          }
        }
      }
      if (!items.length) { alert('발주할 항목이 없습니다.'); setDownloading(false); return; }

      const res = await fetch('/api/generate-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ vendorId: vendor.id, requester, items }),
      });
      if (!res.ok) throw new Error('서버 오류 ' + res.status);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const today = new Date();
      const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
      a.href = url;
      a.download = `${vendor.name}_발주서_${ds}.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);
      addLogEntry({ vendor: vendor.id, tab: 'standard', items });
    } catch (e) {
      alert('다운로드 실패: ' + e.message);
    }
    setDownloading(false);
  };

  const inSupplies  = activeVendorId === 'supplies';
  const activeVendor = (activeVendorId && !inSupplies) ? VENDOR_MAP[activeVendorId] : null;
  const activeSupply = inSupplies && activeSupplyVendorId ? VENDOR_MAP[activeSupplyVendorId] : null;

  // Summary for repmedicare
  const boxesByProduct = Object.fromEntries(CALCULATED_PRODUCTS.map(p => [p.id, calcTotalBoxes(p)]));
  const grandTotal = Object.values(boxesByProduct).reduce((s, b) => s + b, 0);
  const meetsMin = grandTotal >= FREE_SHIP_BOXES;

  if (!hydrated) return null;

  return (
    <>
      <Head>
        <title>통합 발주 관리</title>
        <meta name="description" content="업체별 발주서 통합 관리" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="icon" href="/favicon.ico" />
      </Head>
      <div className={styles.page}>
        <header className={styles.header}>
          <h1 className={styles.title}>
            {activeVendor ? activeVendor.name + ' 발주' : inSupplies ? '소모품 발주' : '발주 관리'}
          </h1>
          <p className={styles.subtitle}>
            {(activeVendor || inSupplies) ? '← 홈으로 돌아가려면 뒤로가기 버튼을 누르세요' : '업체별 발주서 통합 관리'}
          </p>
          {(activeVendor || inSupplies) && (
            <button onClick={() => { setActiveVendorId(null); setActiveSupplyVendorId(null); }} className={styles.resetBtn}>← 업체 목록</button>
          )}
        </header>

        <main className={styles.main}>
          {/* ── HOME: vendor grid ── */}
          {!activeVendor && (
            <>
              <div className={styles.section} style={{ padding: '18px 20px' }}>
                <p style={{ margin: '0 0 14px', fontSize: '0.9rem', color: '#555', fontWeight: 600 }}>업체 선택</p>
                <div className={styles.vendorGrid}>
                  {HOME_VENDORS.map(v => (
                    <VendorCard key={v.id} vendor={v} onClick={() => { setActiveVendorId(v.id); setActiveSupplyVendorId(null); }} />
                  ))}
                </div>
              </div>
              <OrderLog
                log={orderLog}
                onClear={() => {
                  setOrderLog([]);
                  fetch('/api/log', { method: 'DELETE' }).catch(() => {});
                }}
              />
            </>
          )}

          {/* ── VENDOR DETAIL ── */}
          {(activeVendor || inSupplies) && (
            <>
              {/* Requester row */}
              <div className={styles.requesterRow}>
                <span className={styles.requesterLabel}>출고의뢰인</span>
                <input
                  type="text"
                  value={requester}
                  onChange={e => setRequester(e.target.value)}
                  placeholder="이름 입력"
                  className={styles.requesterInput}
                />
              </div>

              {/* ── 소모품 (grouped vendors) ── */}
              {inSupplies && (
                <div className={styles.section}>
                  <div style={{ display: 'flex', gap: 8, padding: '14px 20px 10px', flexWrap: 'wrap' }}>
                    {SUPPLY_VENDORS.map(v => (
                      <button key={v.id} onClick={() => setActiveSupplyVendorId(v.id)}
                        style={{ padding: '9px 18px', borderRadius: 10, border: '2px solid ' + (activeSupplyVendorId === v.id ? v.color : '#dde3f0'), background: activeSupplyVendorId === v.id ? v.color : '#fff', color: activeSupplyVendorId === v.id ? '#fff' : '#333', fontWeight: 700, fontSize: 14, cursor: 'pointer', transition: 'all .15s' }}>
                        {v.name}
                      </button>
                    ))}
                  </div>
                  {activeSupply ? (
                    <>
                      {activeSupply.products.map(prod => (
                        <ManualProductSection
                          key={prod.id}
                          product={prod}
                          orders={(vendorDrafts[activeSupply.id] || {})[prod.id] || {}}
                          onChange={(pid, sid, val) => handleVendorDraftChange(activeSupply.id, pid, sid, val)}
                        />
                      ))}
                      <div className={styles.summary} style={{ background: '#fff', border: '1px solid #dde3f0' }}>
                        <button onClick={() => handleVendorDownload(activeSupply)} disabled={downloading} className={styles.downloadBtn} style={{ margin: 0 }}>
                          {downloading ? '생성 중...' : '📄 발주서 엑셀 다운로드'}
                        </button>
                      </div>
                    </>
                  ) : (
                    <div style={{ padding: '30px 20px', color: '#aaa', textAlign: 'center', fontSize: 14 }}>위에서 업체를 선택하세요</div>
                  )}
                </div>
              )}

              {/* ── Medyssey (history type) ── */}
              {activeVendor?.type === 'history' && (
                <div className={styles.section}>
                  <MedysseyTab
                    adds={medysseyAdds}
                    onAddInst={handleAddInst}
                    onAddSize={handleAddSize}
                    onRemoveSize={handleRemoveSize}
                    cart={medysseyCart}
                    onCartChange={setMedysseyCart}
                    onDownload={handleMedysseyDownload}
                    downloading={medDownloading}
                    requester={requester}
                  />
                </div>
              )}

              {/* ── Standard vendors (calculated + manual) ── */}
              {activeVendor && activeVendor.type !== 'history' && (() => {
                const calcProds = activeVendor.products.filter(p => p.type === 'calculated');
                const manualProds = activeVendor.products.filter(p => p.type === 'manual');
                const isRepmed = activeVendor.id === 'repmedicare';

                return (
                  <>
                    {calcProds.length > 0 && <UploadSection onUpload={handleUpload} />}
                    {calcProds.length > 0 && (
                      <OchaPanel ocha={ocha} onChange={handleOchaChange} onReset={handleOchaReset} />
                    )}
                    {calcProds.map(prod => (
                      <ProductTable
                        key={prod.id}
                        product={prod}
                        stocks={stocks[prod.id]}
                        jeonsan={jeonsan[prod.id]}
                        ocha={ocha[prod.id]}
                        adjustments={adjustments[prod.id]}
                        onStockChange={(sid, val) => handleStockChange(prod.id, sid, val)}
                        onAdjust={(sid, delta) => handleAdjust(prod.id, sid, delta)}
                        onResetAdj={() => resetProductAdj(prod.id)}
                      />
                    ))}
                    {manualProds.map(prod => {
                      const orders = isRepmed
                        ? manualOrders
                        : (vendorDrafts[activeVendor.id] || {});
                      return (
                        <ManualProductSection
                          key={prod.id}
                          product={prod}
                          orders={orders[prod.id] || {}}
                          onChange={isRepmed
                            ? handleManualOrderChange
                            : (pid, sid, val) => handleVendorDraftChange(activeVendor.id, pid, sid, val)}
                        />
                      );
                    })}

                    {/* Summary & download */}
                    {isRepmed ? (
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
                          : <p className={styles.summaryMsg}>⚠️ 현재 {grandTotal}박스 — 무료배송까지 <strong>{FREE_SHIP_BOXES - grandTotal}박스</strong> 부족</p>}
                        <button onClick={() => handleVendorDownload(activeVendor)} disabled={downloading || grandTotal === 0} className={styles.downloadBtn}>
                          {downloading ? '생성 중...' : '📄 발주서 엑셀 다운로드'}
                        </button>
                      </div>
                    ) : (
                      <div className={styles.summary} style={{ background: '#fff', border: '1px solid #dde3f0' }}>
                        <button onClick={() => handleVendorDownload(activeVendor)} disabled={downloading} className={styles.downloadBtn} style={{ margin: 0 }}>
                          {downloading ? '생성 중...' : '📄 발주서 엑셀 다운로드'}
                        </button>
                      </div>
                    )}

                    {isRepmed && (
                      <div className={styles.infoBox}>
                        <strong>계산 기준</strong>
                        <ul>
                          <li>적정재고 = 최근 3개월 월평균 사용량 × 1배수</li>
                          <li>발주수량 = (적정재고 − 창고재고) 30개 단위 올림</li>
                          <li>1팩 = 30개 / 1박스 = 60개 (2팩)</li>
                          <li>팩수가 홀수면 테이블 하단에서 조정 방법을 선택할 수 있습니다</li>
                          <li>무료배송 기준: 합계 6박스 이상</li>
                        </ul>
                      </div>
                    )}
                  </>
                );
              })()}

              <OrderLog
                log={orderLog}
                onClear={() => {
                  setOrderLog([]);
                  fetch('/api/log', { method: 'DELETE' }).catch(() => {});
                }}
              />
            </>
          )}
        </main>
      </div>
    </>
  );
}
