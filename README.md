# 랩밴드 / 메디켓 발주 관리 시스템

Next.js 웹 앱. 랩밴드·메디켓 발주 수량 계산 + 소모품 발주서 자동 생성.
Vercel 서버리스 배포. 공유 로그는 Upstash Redis에 저장.

---

## AI가 코드를 수정하기 전에 반드시 읽을 것

- **`AGENTS.md` 먼저 읽기** — 이 Next.js 버전은 일반 학습 데이터와 다른 breaking changes가 있음.
- **Vercel 서버리스 = 파일 시스템 없음** — Excel 템플릿은 `lib/orderTemplate.js`에 base64로 내장. 템플릿 파일이 바뀌면 Python 스크립트로 재인코딩 후 커밋해야 함.
- **Excel 라이브러리 두 개** — `xlsx`(SheetJS)는 클라이언트에서 파일 읽기 전용. `exceljs`는 서버에서 발주서 생성 전용. 절대 바꾸지 말 것 — `exceljs`만 병합 셀·서식을 유지함.
- **제품 추가는 `lib/products.js` 한 곳만** — `pages/index.js`나 API 수정 불필요.

---

## 파일 구조

```
lib/
  products.js        ← 제품·사이즈 데이터 단일 진실 원천 (제품 추가 시 여기만 수정)
  orderTemplate.js   ← 발주서 Excel 템플릿 base64 인코딩 (렙메디케어 시트)
  redis.js           ← Upstash Redis 클라이언트 (env 미설정 시 graceful no-op)

pages/
  index.js           ← 전체 UI. 모든 컴포넌트가 이 파일 한 곳에 있음
  api/
    generate-order.js        ← 랩밴드/메디켓 발주서 Excel 생성 (POST)
    generate-supply-order.js ← 소모품 발주서 Excel 생성 (POST)
    log.js                   ← 공유 발주 로그 API (GET / POST / DELETE)

styles/
  Home.module.css    ← 전체 스타일
```

---

## UI 구조 (pages/index.js)

페이지는 두 탭으로 구성됨:

### 💊 랩밴드/메디켓 발주 탭
- `UploadSection` — 전산 Excel 업로드(Mediquet·Rapband) → 창고재고 자동 계산
- `OchaPanel` — 오차(창고재고 − 전산재고) 편집 패널 (접을 수 있음)
- `ProductTable` (calculated 제품마다) — 재고·발주수량·팩수·박스 수 표시, 홀수팩 조정 UI
- `ManualProductSection` (manual 제품마다) — 스타킹 등 수기 수량 입력
- 요약 패널 — 총 박스 수, 무료배송 여부, 발주서 Excel 다운로드 버튼

### 📦 소모품 발주 탭
- `SupplyOrderTab` — 출고의뢰인·수주처 명 입력 + 품명/규격/수량 동적 행 추가 (최대 13행)
- 다운로드 시 파일명: `{첫번째품명}_발주서_{날짜}.xlsx`
- 폼 상태는 `localStorage(supplyOrderDraft)`에 자동 저장

### 발주 이력 로그 (두 탭 공통, 항상 하단 표시)
- `OrderLog` — 발주서 다운로드마다 자동 기록
- Upstash Redis(`rapband:orderLog` 키)에 저장 → 모든 사용자가 공유
- 최대 200건, "이력 삭제" 버튼으로 초기화

---

## 제품 추가 방법

**`lib/products.js`의 `PRODUCTS` 배열에 항목만 추가하면 끝.**
UI, API 수정 불필요.

### 계산형 제품 (재고 기반 자동 계산 — Mediquet·Rapband와 같은 형태)

```js
{
  id: 'newproduct',        // 고유 slug, state 키 및 API payload에 사용
  name: '표시이름',
  type: 'calculated',
  color: '#hex',           // 섹션 헤더 색상
  excelName: 'ExcelName', // 발주서 Excel 품명 컬럼에 기입될 텍스트
  hasExcelUpload: true,    // 업로드 섹션에 파일 선택창 표시 여부
  defaultOcha: { M1: 0, M2: 0, L: 0, XL: 0 }, // 오차 기본값
  sizes: [
    {
      id: 'M1',
      label: 'M1 빨강',
      chipColor: '#e74c3c',
      excelSize: 'M1(빨강)',   // 발주서 규격 컬럼에 기입될 텍스트
      avg_monthly: 50.0,
      monthly: { '2026-03': 45, '2026-04': 52, '2026-05': 53 },
      default_stock: 80,
    },
  ],
},
```

### 수기형 제품 (사용자가 수량 직접 입력 — 스타킹과 같은 형태)

```js
{
  id: 'newmanual',
  name: '표시이름',
  type: 'manual',
  color: '#hex',
  excelName: 'ExcelName',
  hasExcelUpload: false,
  sizes: [
    { id: 'L', label: 'L', chipColor: '#hex', excelSize: 'L', default_stock: 0 },
  ],
},
```

---

## 데이터 흐름

```
[랩밴드/메디켓 탭]

Excel 업로드 (클라이언트)
  → xlsx로 각 시트의 D3 셀 읽기 (시트명에 (SIZE) 포함 여부로 매칭)
  → handleUpload: 창고재고 = 전산값 + 오차
  → stocks / jeonsan 상태 업데이트, adjustments 초기화

발주수량 계산 (순수 함수)
  safetyStock = ceil(avg_monthly)   ← 1배수
  shortage    = safetyStock − currentStock
  packs       = ceil(shortage / 30)
  qty         = packs × 30

홀수팩 조정
  → 줄이기/늘리기 클릭 → adjustments[product][size] ±= 1

발주서 Excel 생성
  → POST /api/generate-order  { orders: [{ product, size, qty }] }
  → 서버: 템플릿 로드 → D4=오늘 날짜 → 10-22행 초기화
  → 제품 그룹 사이 빈 행 삽입하며 기입
  → .xlsx 스트리밍 반환

[소모품 탭]

  → POST /api/generate-supply-order { requester, recipient, items }
  → 서버: 템플릿 로드 → D4=날짜, M4=출고의뢰인, D6=수주처 명
  → 품목 행 기입 (그룹 없음)
  → .xlsx 반환, 파일명 = {첫번째품명}_발주서_{날짜}.xlsx

[공유 로그]

  → 다운로드 성공 시 addLogEntry() 호출
  → 로컬 state 즉시 업데이트 (optimistic)
  → POST /api/log → Upstash Redis LPUSH rapband:orderLog
  → 페이지 로드 시 GET /api/log → Redis LRANGE로 전체 로그 복원
```

---

## localStorage 스키마

키: `bandOrderStocks`

```json
{
  "stocks":       { "mediquet": { "M1": 52 }, "rapband": { "M1": 84 } },
  "jeonsan":      { "mediquet": { "M1": null }, "rapband": { "M1": null } },
  "adjustments":  { "mediquet": { "M1": 0 }, "rapband": { "M1": 0 } },
  "ocha":         { "mediquet": { "M1": 22 }, "rapband": { "M1": 3 } },
  "manualOrders": { "stocking": { "L": 0, "XL": 0 } }
}
```

키: `supplyOrderDraft`

```json
{ "requester": "이민재", "recipient": "렙메디케어", "items": [{ "name": "", "spec": "", "qty": 1 }] }
```

---

## Excel 템플릿 (렙메디케어 시트)

| 셀 | 내용 |
|---|---|
| D4 | 출고의뢰일 (오늘 날짜로 자동 설정) |
| M4 | 출고의뢰인 (소모품 탭에서 입력값 기입) |
| A6 | 수주처 명 (레이블) |
| D6 | 수주처 명 값 (소모품 탭에서 입력값 기입) |
| A10:H22 | 발주 품목 행 (최대 13행) |

열 매핑: `A`=번호, `D`=품명, `F`=규격, `H`=수량

서식·병합 셀은 ExcelJS가 보존. 템플릿 수정 시 Python으로 재인코딩:
```python
import base64
with open('template.xlsx', 'rb') as f:
    print(base64.b64encode(f.read()).decode())
# → lib/orderTemplate.js의 TEMPLATE_B64 값 교체
```

---

## 환경변수 (Vercel)

| 변수 | 설명 |
|---|---|
| `UPSTASH_REDIS_REST_URL` | Upstash Redis REST endpoint |
| `UPSTASH_REDIS_REST_TOKEN` | Upstash Redis REST token |

미설정 시 로그 API만 비활성화, 나머지 기능 정상 동작.

---

## 비즈니스 규칙

| 규칙 | 값 |
|---|---|
| 팩 단위 | 30개 |
| 박스 단위 | 60개 (2팩) |
| 무료배송 기준 | 합계 6박스 이상 (calculated 제품 합산) |
| 적정재고 배수 | 월평균 × 1배수 |
| 발주 단위 | 팩 단위 올림 |
| 홀수팩 제약 | 다운로드 전 줄이기/늘리기로 해결 필요 |

---

## 배포

Vercel. `main` 브랜치 push 시 자동 배포.
