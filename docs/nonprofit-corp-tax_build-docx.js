/*
 * docs/nonprofit-corp-tax.html → 비영리 사단법인 세무 실무 노트.docx
 *
 * 워드에서 직접 편집하기 위한 산출물이다. 표·본문·흐름도 전부 워드 네이티브
 * 요소(문단·표)로 만들며 그림은 쓰지 않는다. 그림으로 넣으면 워드에서 글자를
 * 고칠 수 없기 때문이다.
 *
 * 실행:
 *   npm install docx
 *   node docs/nonprofit-corp-tax_build-docx.js
 *
 * 주의 (검증된 제약):
 *   - 표는 columnWidths 합 = 표 width 로 정확히 맞출 것. 안 맞으면 워드에서 깨진다.
 *   - 음영은 ShadingType.CLEAR. SOLID 를 쓰면 칸이 검게 나온다.
 *   - 세로 병합은 위쪽 칸에만 rowSpan 을 주고, 아래 행에서는 그 칸을 아예 생략한다.
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign,
  PageOrientation, PageBreak, HeadingLevel,
} = require('docx');

/* ── 색 (HTML 의 :root 토큰과 동일) ───────────────────────────── */
const INK = '1C2320', INK2 = '46534D', INK3 = '75837C';
const JADE = '2A6350', JADE_SOFT = 'E7F0EB';
const CLAY = 'A85F2B', CLAY_SOFT = 'FAEEE2';
const SEAL = 'A8382E';
const KEY_ROW = 'F7F3E6';          // gold 11% — 강조 행
const RULE = 'D8D4C8';
const FOCUS = 'A32E24', ON_DARK = 'FFFFFF';

const FONT = '맑은 고딕';

/* ── 페이지 (A4) ─────────────────────────────────────────────── */
const A4_W = 11906, A4_H = 16838;   // twips
const MARGIN = 1134;                // 20mm
const W_PORTRAIT = A4_W - MARGIN * 2;    // 9638
const W_LANDSCAPE = A4_H - MARGIN * 2;   // 14570

/* ── 텍스트 헬퍼 : **굵게** 마크업을 TextRun 으로 ────────────── */
function rt(text, opts = {}) {
  const base = { font: FONT, size: opts.size || 19, color: opts.color || INK };
  return String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((seg) => {
    const bold = seg.startsWith('**') && seg.endsWith('**');
    return new TextRun({ ...base, bold: bold || !!opts.bold, text: bold ? seg.slice(2, -2) : seg });
  });
}

const para = (text, opts = {}) => new Paragraph({
  children: rt(text, opts),
  alignment: opts.align,
  spacing: { before: opts.before ?? 0, after: opts.after ?? 100, line: opts.line ?? 260 },
});

const h2 = (text) => new Paragraph({
  children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: INK })],
  spacing: { before: 320, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: JADE } },
});

const h3 = (text) => new Paragraph({
  children: [new TextRun({ text, font: FONT, size: 21, bold: true, color: INK })],
  spacing: { before: 180, after: 60 },
});

const lede = (text) => para(text, { color: INK2, size: 18, after: 160 });
const cite = (text) => para(text, { color: INK3, size: 16, after: 140 });

/* ── 표 헬퍼 ─────────────────────────────────────────────────── */
const BORDER = { style: BorderStyle.SINGLE, size: 2, color: RULE };
const TABLE_BORDERS = {
  top: BORDER, bottom: BORDER, left: BORDER, right: BORDER,
  insideHorizontal: BORDER, insideVertical: BORDER,
};

function cell(text, o = {}) {
  const children = Array.isArray(text)
    ? text
    : [new Paragraph({
        children: rt(text, { size: o.size || 18, color: o.color || INK, bold: o.bold }),
        alignment: o.align,
        spacing: { before: 40, after: 40, line: 240 },
      })];
  return new TableCell({
    children,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    rowSpan: o.rowSpan,
    columnSpan: o.columnSpan,
    verticalAlign: o.valign || VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
  });
}

const headCell = (text, o = {}) =>
  cell(text, { ...o, bold: true, size: 17, color: INK2, fill: 'F1EFE8', align: AlignmentType.CENTER });

function table(columnWidths, rows, width) {
  return new Table({
    columnWidths,
    width: { size: width, type: WidthType.DXA },
    borders: TABLE_BORDERS,
    rows,
  });
}

/* ══ 1. 판정 흐름도 — 워드 네이티브 표 2개 ═══════════════════════
   원본 HTML 의 SVG 흐름도를 표로 옮긴 것. 그림이 아니라 표라서
   워드에서 글자를 그대로 고칠 수 있다.                            */

const FLOW1_COLS = [2200, 1800, 2600, 2600, 2600, 2770];   // 합 14570
const flow1 = table(FLOW1_COLS, [
  new TableRow({
    tableHeader: true,
    children: [headCell('판정'), headCell('갈래'), headCell('이후 절차', { columnSpan: 4 })],
  }),
  new TableRow({
    children: [
      cell([
        new Paragraph({ children: rt('수익사업이 있는가?', { bold: true, size: 20 }), alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        new Paragraph({ children: rt('법 §4③ · 영 §3', { size: 16, color: INK3 }), alignment: AlignmentType.CENTER }),
      ], { rowSpan: 3, fill: 'FFFFFF' }),
      cell('수익사업 없음', { bold: true, color: JADE, align: AlignmentType.CENTER }),
      cell([
        new Paragraph({ children: rt('법인세 신고의무 없음', { bold: true, size: 19 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('회비 · 기부금 · 보조금만 있는 경우', { size: 16, color: INK2 }) }),
      ], { columnSpan: 4 }),
    ],
  }),
  new TableRow({
    children: [
      cell('이자 · 배당소득만', { bold: true, color: JADE, align: AlignmentType.CENTER }),
      cell([
        new Paragraph({ children: rt('분리과세 선택 (법 §62)', { bold: true, size: 19 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('원천징수로 종결 · 신고 생략 가능', { size: 16, color: INK2 }) }),
      ], { columnSpan: 4 }),
    ],
  }),
  new TableRow({
    children: [
      cell('수익사업 있음', { bold: true, color: SEAL, align: AlignmentType.CENTER }),
      cell([
        new Paragraph({ children: rt('사업자등록 20일 이내', { bold: true, size: 18 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('수익사업 개시신고 2개월 이내', { size: 15, color: INK2 }) }),
        new Paragraph({ children: rt('별지 제75호의4 서식', { size: 15, color: INK3 }) }),
      ]),
      cell([
        new Paragraph({ children: rt('구분경리 (법 §113)', { bold: true, size: 18 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('수익 · 비수익 분리 기장', { size: 15, color: INK2 }) }),
        new Paragraph({ children: rt('공통손익은 수입금액 안분', { size: 15, color: INK2 }) }),
      ]),
      cell([
        new Paragraph({ children: rt('3.31 법인세 신고', { bold: true, size: 19, color: ON_DARK }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('지방소득세 4.30까지', { size: 15, color: ON_DARK }) }),
        new Paragraph({ children: rt('중간예납 8.31', { size: 15, color: ON_DARK }) }),
      ], { fill: JADE }),
      cell([
        new Paragraph({ children: rt('고유목적사업준비금 (법 §29)', { bold: true, size: 17 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('이자 · 배당 100% · 그 외 50%', { size: 15, color: INK2 }) }),
        new Paragraph({ children: rt('5년 내 미사용분 익금환입', { size: 15, color: SEAL }) }),
      ]),
    ],
  }),
], W_LANDSCAPE);

const FLOW2_COLS = [2200, 2600, 2800, 3800, 3170];        // 합 14570
const flow2 = table(FLOW2_COLS, [
  new TableRow({
    tableHeader: true,
    children: [headCell('판정'), headCell('갈래'), headCell('이후 절차', { columnSpan: 3 })],
  }),
  new TableRow({
    children: [
      cell([
        new Paragraph({ children: rt('공익사업을 하는가?', { bold: true, size: 20 }), alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        new Paragraph({ children: rt('상증령 §12', { size: 16, color: INK3 }), alignment: AlignmentType.CENTER }),
      ]),
      cell('해당 — 학술 · 장학 · 사회복지 · 의료 등', { bold: true, color: CLAY, align: AlignmentType.CENTER }),
      cell([
        new Paragraph({ children: rt('전용계좌 개설 · 신고', { bold: true, size: 18 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('상증법 §50조의2', { size: 15, color: INK2 }) }),
        new Paragraph({ children: rt('미개설 · 미사용 시 가산세', { size: 15, color: SEAL }) }),
      ]),
      cell([
        new Paragraph({ children: rt('4.30 — 네 가지가 한꺼번에', { bold: true, size: 19, color: ON_DARK }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('출연재산 등 보고서 · 결산서류 등 공시', { size: 15, color: ON_DARK }) }),
        new Paragraph({ children: rt('외부전문가 세무확인 · 의무이행 여부 보고', { size: 15, color: ON_DARK }) }),
        new Paragraph({ children: rt('세무확인 대상 : 총자산 5억 이상 또는 수입금액＋출연재산 3억 이상', { size: 14, color: ON_DARK }) }),
      ], { fill: CLAY }),
      cell([
        new Paragraph({ children: rt('6.30 기부금영수증 발급명세서 제출', { bold: true, size: 18 }), spacing: { after: 30 } }),
        new Paragraph({ children: rt('발급명세 5년 보관 · 전자발급분 제외', { size: 15, color: INK2 }) }),
        new Paragraph({ children: rt('법인세법 §112조의2', { size: 15, color: INK3 }) }),
      ]),
    ],
  }),
], W_LANDSCAPE);

/* ══ 2. 수입 항목별 판정표 (13행 × 4열) ════════════════════════ */

const VERDICT_COLS = [3400, 1100, 1000, 4138];             // 합 9638
const V = { in: SEAL, out: JADE, chk: CLAY, na: INK3 };

const verdictRows = [
  ['정회원 · 단체회원 회비 (반대급부 없음)', ['비수익', V.out], ['해당 없음', V.na], '정관 · 회칙에 회비 대가로 주는 것이 적혀 있는지 확인', false],
  ['회비인데 학술지 제공 · 학술대회 무료참가 등 **급부가 붙는 경우**', ['대가성 검토', V.chk], ['검토', V.chk], '급부의 가치가 회비에 가까울수록 수익사업에 근접', false],
  ['학술대회 · 세미나 **참가비 · 등록비**', ['수익사업', V.in], ['면세', V.out], '주무관청 허가 법인이 실비 수준으로 받을 때 면세. 법인격 없는 임의단체는 과세', true],
  ['학술지 **게재료 · 심사료**', ['수익사업', V.in], ['면세', V.out], '실비 요건 충족 전제', false],
  ['학술지 **광고료**', ['수익사업', V.in], ['면세', V.out], '비영리 출판물에 게재하는 광고', false],
  ['학술대회 **부스 임대료 · 전시 참가비**', ['수익사업', V.in], ['과세', V.in], '연구 관련성 · 실비 요건을 벗어나므로 **부가세까지 붙습니다**. 세금계산서 발급 필요', true],
  ['**수탁 연구용역** (부처 · 공공기관 발주)', ['수익사업', V.in], ['검토', V.chk], '**계약서 형태로 판정합니다.** 용역계약이면 대가관계 있음 → 수익사업', true],
  ['**국고보조금 · 지자체 보조금**', ['비수익', V.out], ['해당 없음', V.na], '교부결정 통지 방식. 위 수탁용역과 **반드시 구분**', true],
  ['기부금 · 후원금 (반대급부 없음)', ['비수익', V.out], ['해당 없음', V.na], '지정 공익법인이면 기부금영수증 발급 → 6.30 명세서 제출로 연결', false],
  ['단행본 · 연구보고서 판매', ['수익사업', V.in], ['검토', V.chk], '실비 초과 여부로 갈림', false],
  ['예치금 · 정기예금 **이자**', ['수익사업', V.in], ['해당 없음', V.na], '고유목적사업준비금 **100% 설정** 가능 → 실질 세부담 없음', false],
  ['사무실 임대 · 강의실 대관', ['수익사업', V.in], ['과세', V.in], '부수적이어도 계속 · 반복이면 수익사업', false],
];

const verdictTable = table(VERDICT_COLS, [
  new TableRow({
    tableHeader: true,
    children: [headCell('수입 항목'), headCell('법인세'), headCell('부가세'), headCell('메모')],
  }),
  ...verdictRows.map(([item, ct, vat, memo, key]) => new TableRow({
    children: [
      cell(item, { fill: key ? KEY_ROW : undefined }),
      cell(ct[0], { color: ct[1], bold: true, align: AlignmentType.CENTER, fill: key ? KEY_ROW : undefined }),
      cell(vat[0], { color: vat[1], bold: true, align: AlignmentType.CENTER, fill: key ? KEY_ROW : undefined }),
      cell(memo, { size: 17, color: INK2, fill: key ? KEY_ROW : undefined }),
    ],
  })),
], W_PORTRAIT);

/* ══ 3. 연간 세무일정 (17행 × 3열) ═════════════════════════════ */

const SCHED_COLS = [1200, 3400, 5038];                     // 합 9638

const schedRows = [
  ['매월 10일', '원천세 신고 · 납부', '발표 사례비 · 원고료 · 심사료는 기타소득(필요경비 60% 의제). 반기납 승인 시 1.10 · 7.10', ''],
  ['1.25', '부가가치세 2기 확정신고', '과세 수익사업이 있을 때만', ''],
  ['2.10', '사업장현황신고', '면세사업만 있는 경우. 부스 임대 등 과세분이 있으면 부가세 신고로 넘어감', ''],
  ['2월 말', '이자 · 배당 · **기타소득 지급명세서**', '학술단체 원천징수의 주력. 연말정산은 2월분 급여 지급 시', 'key'],
  ['3.10', '근로 · 퇴직 · 사업소득 지급명세서', '', ''],
  ['3.31', '[법인세] 법인세 신고 · 납부', '사업연도 종료일부터 3개월. 이자소득만 있고 분리과세를 택하면 생략', 'key'],
  ['2 ~ 3월', '정기 사원총회 결산 승인 · 주무관청 사업실적 및 결산서 제출', '세무 업무는 아니지만 부처 규칙상 종료 후 2 ~ 3개월. **법인세 신고보다 먼저인 경우가 많습니다**', 'muted'],
  ['4.25', '부가가치세 1기 예정신고', '법인사업자', ''],
  ['4.30', '법인지방소득세 신고 · 납부', '법인세 신고기한 종료일이 속하는 달의 말일부터 1개월', ''],
  ['4.30', '[공익법인] 출연재산 등 보고서 · 결산서류 등 공시 · 외부전문가 세무확인 · 의무이행 여부 보고', '모두 사업연도 종료일부터 4개월. **공익법인이면 4월이 3월보다 바쁩니다**', 'key'],
  ['6.30', '[공익법인] 기부금영수증 발급명세서 제출', '사업연도 종료일부터 6개월. 전자기부금영수증 발급분은 제외', 'key'],
  ['7.10', '원천세 반기납 1기분', '반기납 승인 법인', ''],
  ['7.25', '부가가치세 1기 확정신고', '', ''],
  ['8.31', '법인세 중간예납', '직전 사업연도 산출세액이 있는 경우', ''],
  ['10.25', '부가가치세 2기 예정신고', '', ''],
  ['수시', '수익사업 개시 → 사업자등록 **20일 이내** · 수익사업 개시신고 **2개월 이내**', '법인세법 시행규칙 별지 제75호의4. 신규 수임 시 과거 누락분부터 확인', 'key'],
];

const schedTable = table(SCHED_COLS, [
  new TableRow({
    tableHeader: true,
    children: [headCell('시기'), headCell('항목'), headCell('근거 · 비고')],
  }),
  ...schedRows.map(([when, what, note, flag]) => {
    const fill = flag === 'key' ? KEY_ROW : undefined;
    const dim = flag === 'muted' ? INK3 : undefined;
    return new TableRow({
      children: [
        cell(when, { bold: true, color: dim || (flag === 'key' ? SEAL : INK), fill }),
        cell(what, { color: dim || INK, fill }),
        cell(note, { size: 17, color: dim || INK2, fill }),
      ],
    });
  }),
], W_PORTRAIT);

/* ══ 4. 사고가 나는 지점 (8개) ═════════════════════════════════ */

const points = [
  ['“면세단체라 법인세는 없다”',
   '학술단체에서 가장 흔한 오해입니다. 부가세 면세는 실비 · 무상 공급 여부를 보고, 법인세 수익사업은 대가관계를 봅니다. 참가비 · 게재료 · 광고료는 **면세이면서 동시에 수익사업**입니다. 과거 신고 누락이 있으면 여기서 나옵니다.',
   '부가가치세법 §26①18호 · 시행령 §45 ↔ 법인세법 §4③'],
  ['보조금과 수탁용역을 갈라야 합니다',
   '부처 · 공공기관에서 들어온 돈을 전부 보조금으로 처리해 둔 경우가 많습니다. **교부결정 통지**면 비수익, **용역계약**이면 대가관계가 있어 수익사업입니다. 통장이 아니라 계약서철을 봐야 판정이 됩니다.',
   '법인세법 시행령 §3 · 계약 형태로 판정'],
  ['원천징수 건수가 많고 관리가 허술합니다',
   '학술대회 발표 · 토론 사례비, 원고료, 논문 심사료는 대부분 기타소득(필요경비 60% 의제)입니다. 지급 대상이 매번 바뀌어 인적사항이 누락되기 쉽고, 그대로 두면 **지급명세서 제출 불성실 가산세**로 돌아옵니다.',
   '기타소득 지급명세서 2월 말 · 사업소득 3.10'],
  ['부스 임대료는 부가세까지 붙습니다',
   '학술대회 부스 · 전시 참가비, 사무실 · 강의실 대관료는 연구 관련성과 실비 요건을 벗어나 **과세**입니다. 과세분이 있으면 면세사업장 현황신고(2.10)가 아니라 부가세 신고 체계로 넘어갑니다.',
   '부가가치세법 시행령 §45 · 실비 요건'],
  ['구분경리 세팅은 첫 달에',
   '수익 · 비수익의 자산 · 부채 · 손익을 나눠 기장하고, 공통손익은 수입금액 비례로 안분합니다. 학술단체는 인건비 · 임차료 · 사무비가 대부분 공통손익이라 부문코드를 처음에 잡아두지 않으면 연말에 손으로 쪼개게 됩니다.',
   '법인세법 §113 · 시행규칙 §76'],
  ['준비금은 설정보다 사용 관리',
   '예치금 이자는 100% 설정이 되므로 실질 세부담이 없지만, 5년 내 미사용분은 익금환입에 이자상당액까지 붙습니다. 선입선출로 소진되므로 설정연도별 잔액표를 따로 굴려야 합니다.',
   '법인세법 §29'],
  ['학술단체는 공익법인 트랙에 걸립니다',
   '학술 연구 · 발전을 목적으로 하면 상증법상 공익법인에 해당합니다. 전용계좌 개설 · 신고, 4.30 네 가지 보고가 그대로 따라붙고, **전용계좌 미개설 가산세**가 대표적인 사고 유형입니다.',
   '상증세법 시행령 §12 · 상증세법 §50조의2'],
  ['기부금단체 지정 기간은 6년입니다',
   '학술연구단체는 공익법인등 지정 대상입니다. 지정을 받아야 기부금영수증을 발급할 수 있고, **지정기간이 끝나기 전에 재지정 신청**을 해야 발급이 끊기지 않습니다. 지정 공고문에서 기간부터 확인합니다.',
   '법인세법 시행령 §39'],
];

/* ══ 5. 실무서 · 참고자료 ══════════════════════════════════════ */

const refs = [
  ['비영리법인 회계와 세무실무', '',
   '삼일회계법인 비영리전문팀 (변영선 · 정미향) · 삼일인포마인 · 매년 개정 · 정가 10만원대',
   '이 분야 사실상의 표준 실무서. 구분경리 · 고유목적사업준비금 · 공익법인 의무를 한 권에서 연결해 다루는 책이 이것 말고는 마땅치 않습니다. 매년 개정되므로 주문 전에 최신판인지 확인하세요.'],
  ['공익법인 세무안내', '무료',
   '국세청 · 매년 2 ~ 3월 발간 · 국세청 > 국세신고안내 > 법인신고안내 > 공익법인 > 참고자료실',
   '공익법인 의무 전반이 서식 · 기한 · 가산세 표로 정리돼 있어 실무에서 제일 자주 펴게 됩니다. 4월 30일 네 가지를 이 책자 순서대로 따라가면 됩니다.'],
  ['공익법인회계기준 실무지침서', '무료',
   '기획재정부 · 공익법인회계기준(기재부 고시) 해설',
   '계정과목별 회계처리 사례집. 결산서류를 조판할 때 근거로 씁니다.'],
  ['비영리조직회계기준', '무료',
   '한국회계기준원 · 2017년 제정',
   '공익법인회계기준 적용 대상이 아닌 법인의 결산 근거. 주무관청 규칙(사회복지법인 재무 · 회계규칙 등)이 따로 있으면 그쪽이 우선입니다.'],
  ['비영리법인 업무편람 · 업무 매뉴얼', '무료',
   '주무관청 발간 (경기도 · 문화체육관광부 · 법무부 등)',
   '정관 변경 · 총회 · 주무관청 보고 부분. 세무사 업무는 아니지만 결산 일정이 여기에 걸리므로 담당 부처 편람 한 부는 받아둡니다.'],
];

/* ══ 6. 신규 수임 시 받을 자료 ═════════════════════════════════ */

const intake = [
  '**정관** · 회칙 — 회비 대가로 무엇을 주는지',
  '**법인설립허가증** · 등기부등본 — 주무관청 확인',
  '**고유번호증** 또는 사업자등록증',
  '최근 3년 **결산서** 및 법인세 신고서',
  '주무관청 제출 **사업보고서 · 결산보고서**',
  '기부금단체 **지정 공고문** — 지정기간 확인',
  '**총회 · 이사회 회의록**',
  '**통장 목록** — 전용계좌 개설 여부',
  '**연구용역 계약서철** — 보조금과 구분하는 핵심 자료',
  '보조금 **교부결정 통지서** · 정산보고서',
  '**학술대회 결산자료** — 참가비 · 부스 · 광고 수입 내역',
  '**학술지 발간 내역** — 게재료 · 심사료 · 광고료',
  '회비 **수납 내역** · 기부금영수증 발급 이력',
  '**원천징수이행상황신고서** 및 지급명세서 제출 이력',
];

/* ── 문서 조립 ───────────────────────────────────────────────── */

const landscape = {
  properties: {
    page: {
      // docx 는 세로 기준 치수를 받아 landscape 일 때 스스로 뒤집는다.
      // 여기서 미리 뒤집어 넘기면 도로 세로가 되므로 A4 세로 치수를 그대로 준다.
      size: { width: A4_W, height: A4_H, orientation: PageOrientation.LANDSCAPE },
      margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
    },
  },
  children: [
    new Paragraph({
      children: [new TextRun({ text: '실 무 노 트', font: FONT, size: 16, bold: true, color: SEAL })],
      spacing: { after: 60 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: '비영리 사단법인 세무', font: FONT, size: 40, bold: true, color: INK }),
        new TextRun({ text: '   학술단체 · 12월 결산 기준', font: FONT, size: 18, color: INK3 }),
      ],
      spacing: { after: 140 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: FOCUS } },
    }),
    para('학술단체에서 가장 많이 어긋나는 지점은 **부가세 면세와 법인세 비수익사업을 같은 것으로 아는 것**입니다. 학술대회 참가비는 부가세는 면세여도 법인세에서는 수익사업입니다. 여기에 법인세 트랙과 공익법인 트랙이 서로 다른 법에서 나오는 별개의 의무로 얹힙니다.',
      { color: INK2, size: 19, after: 240, before: 120 }),

    h3('트랙 1 — 법인세 · 법인세법'),
    flow1,
    new Paragraph({ text: '', spacing: { after: 200 } }),

    h3('트랙 2 — 공익법인 · 상속세및증여세법'),
    flow2,
    new Paragraph({ text: '', spacing: { after: 160 } }),

    para('두 트랙은 서로 독립이라 각각 판정합니다. 출발점은 **수입 항목별 대가성**입니다. 순수 회비 · 기부금 · 보조금은 수익사업이 아니지만, 학술대회 참가비 · 게재료 · 광고료처럼 **대가가 오가는 수입**은 수익사업입니다. 학술단체는 이 항목들이 거의 항상 있으므로 실무상 **수익사업 있음** 갈래로 갑니다. 공익법인이면 4월 의무가 별도로 남습니다.',
      { color: INK2, size: 17 }),
  ],
};

const portrait = {
  properties: {
    page: {
      size: { width: A4_W, height: A4_H, orientation: PageOrientation.PORTRAIT },
      margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
    },
  },
  children: [
    h2('수입 항목별 판정'),
    lede('학술단체 실무의 8할이 이 표입니다. **법인세와 부가세는 판정 기준이 다릅니다** — 부가세 면세는 “학술 연구와 관련하여 실비 또는 무상으로 공급하는가”를 보고, 법인세 수익사업은 “대가관계가 있는 계속·반복적 사업인가”를 봅니다. 그래서 **면세이면서 동시에 수익사업**인 항목이 대부분입니다.'),
    verdictTable,
    new Paragraph({
      children: [
        new TextRun({ text: '비수익 ', font: FONT, size: 17, bold: true, color: JADE }),
        new TextRun({ text: '법인세 과세 대상 아님    ', font: FONT, size: 17, color: INK2 }),
        new TextRun({ text: '수익사업 · 과세 ', font: FONT, size: 17, bold: true, color: SEAL }),
        new TextRun({ text: '신고 · 납부 대상    ', font: FONT, size: 17, color: INK2 }),
        new TextRun({ text: '검토 ', font: FONT, size: 17, bold: true, color: CLAY }),
        new TextRun({ text: '계약 형태 · 실비 여부로 갈림', font: FONT, size: 17, color: INK2 }),
      ],
      spacing: { before: 120, after: 200 },
    }),

    new Paragraph({ children: [new PageBreak()] }),

    h2('연간 세무일정'),
    lede('12월 결산 기준입니다. [법인세]는 수익사업이 있을 때, [공익법인]은 상증법상 공익법인일 때만 해당합니다. 부가세 · 원천세는 해당 사업이 있을 때만 걸립니다.'),
    schedTable,

    new Paragraph({ children: [new PageBreak()] }),

    h2('사고가 나는 지점'),
    lede('신규 수임 시 이 여덟 가지부터 확인합니다. 앞의 셋이 과거 신고 누락으로 이어지는 지점입니다.'),
    ...points.flatMap(([title, body, source]) => [h3(title), para(body, { color: INK2, size: 18, after: 40 }), cite(source)]),

    new Paragraph({ children: [new PageBreak()] }),

    h2('실무서 · 참고자료'),
    lede('한 권만 고른다면 1번입니다. 2 ~ 5번은 전부 무료 PDF라 함께 내려받아 두면 좋습니다.'),
    ...refs.flatMap(([name, free, meta, why], i) => [
      new Paragraph({
        children: [
          new TextRun({ text: `${i + 1}. `, font: FONT, size: 21, bold: true, color: SEAL }),
          new TextRun({ text: name, font: FONT, size: 21, bold: true, color: INK }),
          ...(free ? [new TextRun({ text: `  [${free}]`, font: FONT, size: 16, bold: true, color: JADE })] : []),
        ],
        spacing: { before: 180, after: 50 },
      }),
      para(meta, { color: INK3, size: 16, after: 50 }),
      para(why, { color: INK2, size: 18, after: 120 }),
    ]),

    new Paragraph({ children: [new PageBreak()] }),

    h2('신규 수임 시 받을 자료'),
    lede('이 목록이 채워져야 위의 두 트랙 판정이 끝납니다.'),
    ...intake.map((item) => new Paragraph({
      children: rt(item, { size: 18, color: INK }),
      bullet: { level: 0 },
      spacing: { after: 70, line: 250 },
    })),

    new Paragraph({
      children: rt('학술단체(학회 · 포럼) 기준 실무 참고용 요약입니다. 부가세 면세 판정은 “실비 또는 무상” 요건의 사실판단이 들어가므로 개별 항목은 국세청 예규로 확인이 필요하고, 주무관청 보고 기한은 소관 부처의 비영리법인 설립 · 감독 규칙마다 다릅니다. 적용 전에 해당 연도 국세청 공익법인 세무안내와 소관 부처 규칙으로 확인하십시오.',
        { size: 16, color: INK3 }),
      spacing: { before: 400 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: RULE } },
    }),
  ],
};

const doc = new Document({
  creator: '실무 노트',
  title: '비영리 사단법인 세무 실무 노트',
  description: '학술단체 · 12월 결산 기준',
  styles: { default: { document: { run: { font: FONT, size: 19, color: INK } } } },
  sections: [landscape, portrait],
});

const out = process.argv[2] || path.join(__dirname, '비영리 사단법인 세무 실무 노트.docx');
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});
